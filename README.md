# Requirements V1

Following wage rules in client's company, W1, W2, W3... as the dollars per hour, and W1 < W2 < W3 < W4 <... so on:

1. Employer's salary will be calculate under week cycles.
2. All working hours during ordinary opening hours from 7 am to 7 pm will be paid W1.
3. All working hours out of ordinary opening hours which is 0 am to 7 am, 7 pm to 24pm will be paid W2.
4. If staffs work in the weekend, then apply W3 for ordinary opening hours, W4 for out of ordinary opening hours.
5. If a staff worked over 12 hours in one day, the work over 12 hours will be paid W5.
6. If a staff has already work over 38 hours this week, the work over 38 hours will be paid W6.
7. If a staff has already work over 38 hours this week, the work over 40 hours will be paid W7.
8. If a working hours can apply multiple rules, higher salary rate will always override.

# Requirements V2

There is new requirements that needs to adjust current algorithm:

1. The client hope to handle the overtime across the midnight, those shift logs will be record like this:

```
Start Time |  Finish Time
7 pm  |  2 am
```

2. The wage calculation will follow the rule, that if it apply overtime salary depends on the shift, if it apply regular salary depends on the day, for an example: George work from Sunday 7pm, to Monday 2am. For the two hours during the Monday midnight:
   1. If it reach the overtime threshold(daily and weekly), this two hours must be paid as the daily overtime for Sunday or weekly overtime for last week.
   2. However, if this two hours doesn't match with any daily/weekly overtime condition, it needs to be paid as regular wage at Monday's rate(Work days early overtime). But this two hours still be count as total working hours on Sunday(last week).
