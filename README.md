# Requirements V1

The client’s company applies wage rules W1, W2, W3, … (in dollars per hour), where **W1 < W2 < W3 < W4 < …**:

1. Salaries are calculated on a weekly cycle.
2. All working hours during **ordinary opening hours (07:00–19:00)** are paid at **W1**.
3. All working hours outside ordinary opening hours (**00:00–07:00, 19:00–24:00**) are paid at **W2**.
4. For weekend work (Saturday and Sunday):
   - **W3** applies during ordinary opening hours.
   - **W4** applies outside ordinary opening hours.
5. If a staff member works more than **12 hours in a single day**, all hours beyond 12 are paid at **W5**.
6. If a staff member works more than **38 hours in a week**, all hours beyond 38 are paid at **W6**.
7. If a staff member works more than **40 hours in a week**, all hours beyond 40 are paid at **W7**.
8. When multiple rules apply to the same hours, the **higher wage rate overrides** the others.

---

# Requirements V2

The client has requested updates to handle overtime across midnight shifts:

1. Shifts may cross midnight and be recorded as follows:

```
Start Time |  Finish Time
7 pm  |  2 am
```

2. Wage calculation rules:

- Overtime depends on the **shift day**.
- Regular pay depends on the **calendar day**.

**Example:** George works from **Sunday 19:00 to Monday 02:00**.

- For the two hours after midnight (on Monday):
  1. If overtime thresholds (daily or weekly) are exceeded, those two hours are paid as **Sunday’s overtime** (daily overtime or weekly overtime for the previous week).
  2. If no overtime thresholds are triggered, the hours are paid at **Monday’s regular rate** (weekday early hours).
  3. Regardless, these two hours still **count towards Sunday’s total hours** (for weekly calculations).

---

# Requirements V3

Further clarifications from the client:

1. **Salary calculation features are no longer required** — disable or remove this functionality.
2. **Saturday and Sunday share the same wage rate**, so weekend hours do not need to be split.
3. The client prefers to **eliminate records weekly** rather than storing historical data. Date queries are therefore unnecessary.
4. Time split sheet shall be **rendered one by one** by all staff members in the record, so staff name queries are unnecessary.
