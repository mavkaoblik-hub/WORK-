#!/usr/bin/env python3
"""Generate fictional deterministic transport fixture, never customer records."""
import argparse
import json
import random
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

SEED = 10142026
TARGET_INCOME_CENTS = 14998791
TARGET_EXPENSE_CENTS = 4821173

def iso(day):
    return day.isoformat() if day else None

def currency(amount):
    return float(Decimal(amount).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def distribute(target, count, rng, variation):
    base=target//count
    shifts=[rng.randint(-variation, variation) for _ in range(count)]
    bias=sum(shifts)
    result=[base + x for x in shifts]
    result[-1]+=target - sum(result)
    assert sum(result) == target and min(result)>0
    return result

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--routes",required=True)
    parser.add_argument("--out",required=True)
    args=parser.parse_args()
    r=json.loads(Path(args.routes).read_text(encoding="utf-8-sig"))
    assert isinstance(r,list) and len(r)==114
    route_ids=[str(x["id"]) for x in r]
    rng=random.Random(SEED)
    surnames=["Тестовий","Прикладний","Умовний","Навчальний","Демо","Зразковий","Макетний"]
    names=["Андрій","Олена","Петро","Марія","Іван","Наталія","Максим"]
    drivers=[{"id":f"Водій_{i+1:02d}","surname":surnames[i],"name":names[i],
              "car":f"DEMO-CAR-{i+1:02d}"} for i in range(7)]
    vehicles=[{"id":f"DEMO-CAR-{i+1:02d}","brand":"ВИГАДАНЕ АВТО",
               "model":f"Навчальний {i+1}","plate":f"TEST{i+1:04d}",
               "fuel":"Diesel","consumption_l_100km":27.1+(i%4),
               "price_pln_l":6.05+(i%3)*0.12,"tank_l":650} for i in range(9)]
    revenues=distribute(TARGET_INCOME_CENTS,70,rng,37000)
    costs=distribute(TARGET_EXPENSE_CENTS,259,rng,6300)
    trips=[]
    for i in range(75):
        done=i<70
        active=70<=i<73
        if done:
            start=date(2026,8,1)+timedelta(days=int(i*1.08))
            end=start+timedelta(days=2+(i%3))
        elif active:
            start=[date(2026,10,1),date(2026,10,9),date(2026,10,16)][i-70]
            end=None
        else:
            start=[date(2026,10,28),date(2026,11,2)][i-73]
            end=None
        km=580+(i*41)%1750
        odo=125000+i*2700
        rev=currency(Decimal(revenues[i])/100) if done else (1550.00 if active else 0.00)
        paid=currency(Decimal(revenues[i])/100 * Decimal("0.65")) if done else 0
        # Deliberately one missing odometer update, one long-running trip.
        now_odo=odo+int(km*0.6) if active and i!=70 else None
        trip={"id":f"DEMO-2026-{i+1:04d}",
              "created":iso(start-timedelta(days=2)), "driver":drivers[i%7]["id"],
              "car":vehicles[i%9]["id"],"route":route_ids[i%len(route_ids)],
              "customer":f"Вигаданий замовник {1+i%12}",
              "plan_start":iso(start),"plan_end":iso(start+timedelta(days=4)),
              "status":"Завершено" if done else ("У рейсі" if active else "Заплановано"),
              "actual_start":iso(start) if not (i>=73) else None,
              "actual_end":iso(end),"odometer_start":odo if not (i>=73) else None,
              "odometer_end":odo+km if done else None,
              "km_actual":km if done else None,
              "revenue_usd":rev,"paid_usd":paid,
              "km_plan":km+int(km*.05),"fuel_liters_plan":round(km*0.29,2),
              "fuel_cost_pln_plan":round(km*0.29*6.19,2),
              "memo":"Лише синтетичні дані для перевірки v10.14",
              "finished_at":iso(end),"driving_h":round(km/68,1) if done else None,
              "waiting_h":i%8 if done else None,"lost_h":i%3 if done else None,
              "odometer_now":now_odo,"odometer_updated":iso(date(2026,10,22)) if now_odo else None}
        trips.append(trip)
    expenses=[]
    categories=["Пальне","Паркінг","Дорожні збори","Сервіс","Інші"]
    codes=["PLN","USD","EUR","UAH"]
    per_unit={"USD":4.02,"EUR":4.39,"UAH":0.095}
    for i,cents in enumerate(costs):
        # 238 receipts on completed trips, 21 receipts on the active trips.
        trip=(i*11)%70 if i<238 else 70+(i-238)%3
        t=trips[trip]
        fx=codes[i%4]
        usd=Decimal(cents)/100
        if fx=="USD": native=usd
        elif fx=="PLN": native=usd*Decimal("4.02")
        elif fx=="EUR": native=usd*Decimal("4.02")/Decimal("4.39")
        else: native=usd*Decimal("4.02")/Decimal("0.095")
        d=date.fromisoformat(t["actual_end"] or t["actual_start"])
        if d>date(2026,10,22):d=date(2026,10,22)
        expenses.append({"created":iso(d),"trip":t["id"],"category":categories[i%len(categories)],
             "parking":"DEMO-PARK-01" if i%5==1 else None,
             "amount":currency(native),"currency":fx,
             "fx_pln_per_unit":1 if fx=="PLN" else per_unit[fx],
             "usd":currency(usd),"date":iso(d),"note":f"Вигадана квитанція №{i+1:03d}"})
    fixture={"kind":"TRANSPORT_DEMO_SEED_V1","fictional":True,
             "as_of":"2026-10-22",
             "fx_pln_per_unit":per_unit, "vehicles":vehicles,"drivers":drivers,
             "trips":trips,"expenses":expenses}
    assert len(fixture["drivers"])==7 and len(fixture["vehicles"])==9
    assert len(trips)==75 and len(expenses)==259
    assert sum(int(round(float(t["revenue_usd"])*100)) for t in trips if t["status"]=="Завершено")==TARGET_INCOME_CENTS
    assert sum(int(round(x["usd"]*100)) for x in expenses)==TARGET_EXPENSE_CENTS
    path=Path(args.out)
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(fixture,ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"Generated: {path} (7 drivers, 9 vehicles, 70 complete, 3 active, 2 planned, 259 receipts)")
if __name__=="__main__":
    main()
