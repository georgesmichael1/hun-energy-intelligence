"""Isolated price XML parser. Offline only; no fetching or publication.

Existing ENTSO-E adapters are unchanged. Reuse rights for price outputs must
be settled separately before any automated production publication.
"""
import datetime as dt
import math
import xml.etree.ElementTree as ET


def parse_prices(xml, expected_eic):
    root=ET.fromstring(xml)
    def tag(e): return e.tag.split('}')[-1]
    def first(e,key): return next((v.text for v in e.iter() if tag(v)==key),None)
    def timestamp(value): return dt.datetime.fromisoformat(value.replace('Z','+00:00'))
    rows=[]
    for ts in root.iter():
        if tag(ts)!='TimeSeries': continue
        if first(ts,'in_Domain.mRID')!=expected_eic or first(ts,'out_Domain.mRID')!=expected_eic:
            raise ValueError('Wrong bidding zone')
        if first(ts,'currency_Unit.name')!='EUR' or first(ts,'price_Measure_Unit.name')!='MWH':
            raise ValueError('Wrong price unit')
        curve=first(ts,'curveType')
        if curve not in ('A01','A03'): raise ValueError('Unsupported curve type')
        for p in ts:
            if tag(p)!='Period': continue
            start,end=timestamp(first(p,'start')),timestamp(first(p,'end'))
            resolution=first(p,'resolution')
            if resolution not in ('PT15M','PT60M'):raise ValueError('Unsupported resolution')
            seconds=900 if resolution=='PT15M' else 3600
            count=int((end-start).total_seconds()/seconds)
            if count<=0 or start+dt.timedelta(seconds=count*seconds)!=end:raise ValueError('Bad period')
            points=[]
            for point in p:
                if tag(point)!='Point':continue
                position=int(first(point,'position'));price=float(first(point,'price.amount'))
                if not math.isfinite(price):raise ValueError('Invalid price')
                points.append((position,price))
            positions=[x[0] for x in points]
            if not points or positions!=sorted(set(positions)) or positions[0]!=1 or positions[-1]>count:
                raise ValueError('Invalid positions')
            if curve=='A01' and positions!=list(range(1,count+1)):raise ValueError('Incomplete sequential curve')
            for index,(position,price) in enumerate(points):
                stop=points[index+1][0] if curve=='A03' and index+1<len(points) else count+1 if curve=='A03' else position+1
                for k in range(position,stop):
                    a=start+dt.timedelta(seconds=(k-1)*seconds)
                    rows.append({'start':a,'end':a+dt.timedelta(seconds=seconds),'price':price})
    seen=set()
    for row in rows:
        if row['start'] in seen:raise ValueError('Overlapping series')
        seen.add(row['start'])
    if not rows:raise ValueError('No price series returned')
    return sorted(rows,key=lambda r:r['start'])
