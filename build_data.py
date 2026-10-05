#!/usr/bin/env python3
"""Чете Excel файла (Zastrahovki8 Agent.xlsx) и генерира data.json за приложението.
Структура на Excel (Sheet1): месечни блокове, редове = полици.
Колони: 0 Клиент | 1 Марка и модел | 2 Рег.номер | 3 Сума полица | 4 Полица номер
        | 5 Поредна вноска | 6 Валидна до | 7 Сума клиент | 8 2% ДЗП | 9 Вноска ГФ/ОФ
        | 10 Комисион | 11 Задържан комисион | 12 Сума отчет | 13 Печалба | ...
"""
import json, sys, os, re
from datetime import datetime, date

try:
    from openpyxl import load_workbook
except ImportError:
    print("Нужен е openpyxl: pip install openpyxl"); sys.exit(1)

EXCEL = os.path.expanduser('~/Documents/Zastrahovki8 Agent.xlsx')
OUT = os.path.expanduser('~/zastrahovki-app/data.json')

MONTHS = ['януари','февруари','март','април','май','юни','юли','август',
          'септември','октомври','ноември','декември']

def translit(s):
    TR = {'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ж':'zh','з':'z','и':'i','й':'y',
          'к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r','с':'s','т':'t','у':'u',
          'ф':'f','х':'h','ц':'ts','ч':'ch','ш':'sh','щ':'sht','ъ':'a','ь':'','ю':'yu','я':'ya',
          'А':'A','Б':'B','В':'V','Г':'G','Д':'D','Е':'E','Ж':'Zh','З':'Z','И':'I','Й':'Y',
          'К':'K','Л':'L','М':'M','Н':'N','О':'O','П':'P','Р':'R','С':'S','Т':'T','У':'U',
          'Ф':'F','Х':'H','Ц':'Ts','Ч':'Ch','Ш':'Sh','Щ':'Sht','Ъ':'A','Ь':'','Ю':'Yu','Я':'Ya'}
    return ''.join(TR.get(ch, ch) for ch in s)

def clean_name(s):
    """Нормализира име: маха ГО/GO, транслитерира, маха излишни интервали."""
    if s is None: return None
    n = str(s).strip()
    if not n: return None
    low = n.lower()
    if low in MONTHS or low in ('имущество','оръжия'):
        return None
    # махни ГО/GO суфикс
    parts = n.split()
    while parts and parts[-1].upper() in ('ГО','GO','ГO','GО'):
        parts.pop()
    n = ' '.join(parts).strip()
    n = translit(n)
    n = re.sub(r'\s+', ' ', n).strip()
    return n or None

def to_date(v):
    if v is None: return None
    if hasattr(v, 'strftime'):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    # dd.mm.yyyy или yyyy-mm-dd
    m = re.match(r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})', s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100: y += 2000
        return f'{y:04d}-{mo:02d}-{d:02d}'
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m: return s[:10]
    return None

def minus_months(iso, months):
    """Връща дата (iso) N месеца назад."""
    from datetime import date
    y, mo, d = int(iso[0:4]), int(iso[5:7]), int(iso[8:10])
    mo2 = mo - months
    y2 = y
    while mo2 <= 0:
        mo2 += 12; y2 -= 1
    # коригирай деня ако месецът е по-кратък
    import calendar
    last = calendar.monthrange(y2, mo2)[1]
    d2 = min(d, last)
    return f'{y2:04d}-{mo2:02d}-{d2:02d}'

def add_months(iso, months):
    """Връща дата (iso) N месеца напред."""
    import calendar
    y, mo, d = int(iso[0:4]), int(iso[5:7]), int(iso[8:10])
    mo2 = mo + months
    y2 = y
    while mo2 > 12:
        mo2 -= 12; y2 += 1
    last = calendar.monthrange(y2, mo2)[1]
    return f'{y2:04d}-{mo2:02d}-{min(d,last):02d}'

def gen_installments(polica_id, enddate, vnoska):
    """Връща ВСИЧКИ вноски (1..N) за полицата, за да показва и предстоящите.
    Логика: крайна дата → начало = край − 1 година.
    Вноска K от N: падеж = начало + (K−1)×3 месеца.
    '1/1' → еднократна (без вноски)."""
    if not vnoska or '/' not in str(vnoska):
        return []
    try:
        numpart, denpart = str(vnoska).split('/')
        den = int(denpart)
    except Exception:
        return []
    if den <= 1:
        return []  # 1/1 = еднократно
    if not enddate:
        return []
    start = minus_months(enddate, 12)
    out = []
    for k in range(1, den + 1):
        due = add_months(start, (k - 1) * 3)
        out.append({'no': k, 'of': den, 'dueDate': due})
    return out

def dedup_policies(policies):
    """Премахва дублирани полици по (кола + крайна дата + тип),
    като запазва записа с най-много информация (номер на полица/вноска)."""
    best = {}
    order = []
    for p in policies:
        key = (p.get('carId'), p.get('endDate'), p.get('type'))
        score = (1 if p.get('polica') else 0) + (1 if p.get('vnoska') else 0) + (1 if p.get('suma_client') else 0)
        if key not in best:
            best[key] = (score, p); order.append(key)
        else:
            if score > best[key][0]:
                best[key] = (score, p)
    out = [best[k][1] for k in order]
    for i, p in enumerate(out):
        p['id'] = 'p' + str(i)
    return out

def dedup_events(events):
    """Премахва дублирани вноски: едно събитие за (кола + тип + номер вноска + падеж)."""
    seen = {}
    out = []
    for e in events:
        key = (e.get('carId'), e.get('type'), e.get('no'), e.get('date'))
        if key in seen:
            # поднови съществуващото с по-голяма сума/по-добра информация
            continue
        seen[key] = e
        out.append(e)
    # преномерирай id-тата
    for i, e in enumerate(out):
        e['id'] = 'i' + str(i)
    return out

def main():
    wb = load_workbook(EXCEL, data_only=True)
    ws = wb['Sheet1']

    clients = {}   # name -> {id,name,phone,notes}
    cars = {}      # plate -> {id, clientId, make, model, year, plate, vin, notes, address}
    policies = []
    events = []
    cseq = 0; carseq = 0; pseq = 0
    month = None

    for row in ws.iter_rows(min_row=2, max_col=15):
        v = [c.value for c in row]
        name = v[0]
        if name is None: continue
        name_s = str(name).strip()
        if not name_s: continue
        if name_s.lower() in MONTHS:
            month = name_s; continue
        cname = clean_name(name_s)
        if cname is None: continue

        # клиент
        if cname not in clients:
            clients[cname] = {'id': 'c'+str(cseq), 'name': cname, 'phone': '', 'notes': ''}
            cseq += 1
        cid = clients[cname]['id']

        # кола по рег.номер
        plate = str(v[2]).strip() if v[2] else ''
        make_model = str(v[1]).strip() if v[1] else ''
        parts = make_model.split(' ', 1)
        make = parts[0] if parts else ''
        model = parts[1] if len(parts) > 1 else ''
        if plate:
            if plate not in cars:
                cars[plate] = {'id':'car'+str(carseq), 'clientId':cid, 'make':make, 'model':model,
                               'year':'', 'plate':plate, 'vin':'', 'notes':'', 'address':''}
                carseq += 1
            carid = cars[plate]['id']
        else:
            # без рег.номер — създай кола без номер (рядко)
            key = make_model+'|'+cname
            if key not in cars:
                cars[key] = {'id':'car'+str(carseq), 'clientId':cid, 'make':make, 'model':model,
                             'year':'', 'plate':'', 'vin':'', 'notes':'', 'address':''}
                carseq += 1
            carid = cars[key]['id']

        # полица
        polno = str(v[4]).strip() if v[4] else ''
        endd = to_date(v[6])
        vnoska = str(v[5]).strip() if v[5] else ''
        suma = v[3] if v[3] is not None else ''
        suma_client = v[7] if v[7] is not None else ''
        profit = v[13] if len(v) > 13 and v[13] is not None else ''
        commission = v[10] if len(v) > 10 and v[10] is not None else ''

        ptype = 'ГО' if polno.startswith('BG/') else ('КАСКО' if polno else '')
        pid = 'p'+str(pseq)
        policies.append({
            'id':pid, 'clientId':cid, 'carId':carid, 'type':ptype,
            'polica':polno, 'endDate':endd or '', 'vnoska':vnoska,
            'suma':str(suma), 'suma_client':str(suma_client),
            'profit':str(profit), 'commission':str(commission), 'month':month or ''
        })
        # вноската за ТОЗИ ред/месец → събитие (от суровия ред, преди дедуп на полиците)
        for inst in gen_installments(pid, endd, vnoska):
            events.append({
                'id':'i'+str(len(events)), 'policyId':pid, 'clientId':cid, 'carId':carid,
                'type':'Вноска', 'no':inst['no'], 'of':inst['of'],
                'date':inst['dueDate'], 'note':'',
                'amount':str(suma_client) if suma_client not in (None,'') else str(suma)
            })
        pseq += 1

    # дедупликирай полиците за досиетата (маха дублирани редове от Excel)
    policies = dedup_policies(policies)

    data = {
        'clients': list(clients.values()),
        'cars': list(cars.values()),
        'policies': policies,
        'events': dedup_events(events)
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(data, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('data.json записан:', OUT)
    print('Клиенти:', len(data['clients']), '| Коли:', len(data['cars']), '| Полици:', len(data['policies']))

if __name__ == '__main__':
    main()
