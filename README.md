# Meeting-Action-Management-System
Meeting & Action Management Tracker

An Excel tracker that classifies meeting actions automatically and shows what needs attention, with a summary report and an interactive list of urgent actions.

<img width="740" height="184" alt="Screenshot 2026-11-21 150047" src="https://github.com/user-attachments/assets/150f4d2b-bf84-4b1c-a46a-bcae5f15eb94" />

<img width="884" height="411" alt="Screenshot 2026-11-20 202629" src="https://github.com/user-attachments/assets/dd7043c2-01b5-4958-8ef3-408a0a1df286" />

<img width="1058" height="320" alt="Screenshot 2026-11-21 152111" src="https://github.com/user-attachments/assets/b74fb296-70c0-4ba1-8f02-d1b85dd015a3" />

[Download the workbook](Meeting-Action-System.xlsx)

What it does

- Tracks actions with owner, priority, dependency, due date and last follow-up date.
- Labels every action's health and follow-up urgency with formulas.
- Summarises results in a report with an overview, needs-attention counts and a chart.
- Lists actions by priority, with a slicer to filter by action health.

Formulas

Action Health labels each action as Completed, Not Known, Overdue, Due Soon (within 5 days) or Up coming.

```excel:

=IFS([@Status]="Completed","Completed",
     [@[Due Date]]="","Not Known",
     [@[Due Date]]<DATE(2026,11,1),"Overdue",
     [@[Due Date]]-DATE(2026,11,1)<=5,"Due Soon",
     TRUE,"Up coming")

Follow-up Urgency flags actions as Urgent, Check or Not Urgent.

```excel
=IFS([@Status]="Completed","Completed",
     OR([@[Action Health]]="Due Soon",[@[Action Health]]="Overdue"),"Urgent",
     OR([@[Last Follow-up]]="",[@[Action Health]]="Not Known"),"Check",
     DATE(2026,11,1)-[@[Last Follow-up]]>7,"Urgent",
     TRUE,"Not Urgent")
```

The report date is fixed at 1 Nov 2026, so results always match the screenshots.

Skills: Excel Tables, IFS, OR, DATE, PivotTables, slicers, charts
