<img src="perpop.png" alt="Andel Projects Limited" width="400">

# CP Drag VBA Anomalies Analysis

Here's a quick overview of how I set up the **Group view in MS Project** to make analyzing the anomalies easier.

## Group View Setup in MS Project

I configured the Group view in MS Project to highlight and organize tasks for simpler analysis. This setup allowed me to focus on specific tasks, zero in on potential anomalies, and systematically address any issues. The process was pretty straightforward:

1. I grouped tasks based on key criteria relevant to CP Drag calculations (e.g., by phase, dependency, or other factors contributing to drag).
2. Filters were applied to zoom in on tasks where anomalies were more likely to show up.

## Process of Identifying Anomalies

The method I used was a bit manual but effective:

1. **Task Reduction:** I started by progressively reducing the task duration or scope for tasks suspected to cause the anomaly.
2. **Observation:** Each reduction was followed by observing the CP Drag values. When the anomaly became evident,I knew we were on the right track.
3. **Investigation:** From there, I drilled down to figure out exactly what was causing the anomaly. This involved a mix of:
   - Reviewing the macro logic in detail.
   - Cross-checking how changes impacted dependencies and drag values.
   - Testing edge cases to confirm the root cause.

## Outcome

The investigation led to identifying the exact cause of the anomalies, which I've now addressed in the latest version of the macro.

Not surprisingly Stephen Deveaux had already predicted the probable cause LOL
Visuals of investigation [here](CP_Drag_forensic_Analysis.pdf)



Cheers,  
[Andy Forrester]

