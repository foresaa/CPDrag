# Project Schedule Quality Checks: Authorities, Guidance, and VBA Macros
<img src="perpop.png" alt="Andel Projects Limited" width="400">

This document summarizes 
### - **key project scheduling authorities**, 
### - provides their **official URLs**, 
### - outlines the **guidance** they recommend for **maximizing project schedule quality**, 
### - and includes **VBA macros** you can use in Microsoft Project to automate common schedule quality checks.

---

## 1. Authorities and Their Guidance

Below is a quick reference table of recognized project scheduling authorities, their websites, and the core scheduling guidance they offer.

| **Authority**                                                                | **Website**                                    | **Key Guidance**                                                                                                                                                                                                                                                  |
| :--------------------------------------------------------------------------- | :--------------------------------------------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **PMI (Project Management Institute)**                                       | [https://www.pmi.org](https://www.pmi.org)     | Publishes the **PMBOK® Guide**, outlining best practices for schedule planning and control. Emphasizes creating a fully networked schedule, baseline integrity, and continual monitoring of schedule performance.                                                 |
| **DCMA (Defense Contract Management Agency)**                                | [https://www.dcma.mil](https://www.dcma.mil)   | Known for the **DCMA 14-Point Assessment**, focusing on a set of metrics (e.g., tasks missing predecessors, negative float, excessive leads/lags) to evaluate schedule health and integrity.                                                                      |
| **GAO (U.S. Government Accountability Office)**                              | [https://www.gao.gov](https://www.gao.gov)     | Publishes the **GAO Schedule Assessment Guide**, which describes ten best practices for a high-quality, reliable schedule: comprehensive, well-constructed, credible, and controlled.                                                                             |
| **AACE International (Association for the Advancement of Cost Engineering)** | [https://web.aacei.org](https://web.aacei.org) | Offers **Recommended Practices** for planning, scheduling, and cost engineering. Stresses avoiding open-ended tasks, setting realistic durations, ensuring resource feasibility, and controlling baseline changes.                                                |
| **ISO (International Organization for Standardization)**                     | [https://www.iso.org](https://www.iso.org)     | Standards such as **ISO 21500** provide a high-level project management framework. While not prescriptive in specific checks, they reinforce the importance of consistent processes, logical networks, baselines, and clear dependencies for schedule management. |

---

## 2. Guidance for Project Quality Checks

Although each authority frames its guidelines slightly differently, common **themes** include:

1. **Complete Logic / No Open-Ended Tasks**

   - Every task should have at least one predecessor (except the first) and one successor (except the last).
   - Prevents “dangling” activities that artificially constrain or inflate float.

2. **Constraint Management**

   - Avoid or minimize “hard constraints” (e.g., Must Finish On) that reduce flexibility.
   - Prefer soft constraints or deadlines for target dates.

3. **Float Analysis**

   - Negative float indicates behind-schedule conditions or infeasible logic.
   - High float could suggest broken links or tasks not driven by the network.

4. **Realistic Durations & Long-Task Checks**

   - Break tasks down so each is short enough to be accurately tracked.
   - Very long tasks can hide detail and obscure critical path analysis.

5. **Resource Assignments & Overallocation**

   - Each non-summary task typically needs assigned resources.
   - Overallocations must be addressed (e.g., leveling, reassigning, or adjusting timelines).

6. **Baseline & Performance Tracking**

   - Setting and maintaining a baseline allows you to measure variance.
   - Schedules without baselines lack a frame of reference for progress.

7. **Leads/Lags & Dependency Types**

   - Avoid excessive lead/lag durations.
   - Use standard Finish-to-Start dependencies unless there’s a valid reason otherwise.

8. **Status Date & Progress**
   - Keep the schedule’s “status date” current.
   - Out-of-sequence progress (task starts before a predecessor is finished) may signal errors in updates.

---

## 3. VBA Macros for Common Schedule Checks

Below you will find individual VBA macros that automate **key** schedule health checks in Microsoft Project.

> **How to Use**
>
> 1. Open MS Project.
> 2. Press **ALT+F11** to open the VBA Editor.
> 3. Go to **Insert → Module** and paste the code snippets.
> 4. Run each macro from **Developer → Macros** or directly from the VBA Editor.

---

### 3.1 Tasks Missing Predecessors or Successors

Checks for tasks (excluding summaries) that have **no predecessors** or **no successors**, which can create open-ended logic.

```vb
Sub CheckMissingDependencies()
    Dim t As Task
    Dim missingDeps As String
    missingDeps = "Tasks missing dependencies:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.Predecessors = "" Then
                    missingDeps = missingDeps & " - " & t.ID & " : " & t.Name & " (No Predecessor)" & vbCrLf
                End If
                If t.Successors = "" Then
                    missingDeps = missingDeps & " - " & t.ID & " : " & t.Name & " (No Successor)" & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox missingDeps, vbInformation, "Missing Dependencies Report"
End Sub
```

### 3.2 Tasks with Constraints

Flags tasks using hard constraints (e.g., Must Finish On, Must Start On) instead of the default “As Soon As Possible.”

```vb

Sub CheckConstraints()
    Dim t As Task
    Dim constraintReport As String
    constraintReport = "Tasks with constraints:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                ' Constraint Type 0 = As Soon As Possible (pjASAP)
                If t.ConstraintType <> pjASAP Then
                    constraintReport = constraintReport & _
                        " - " & t.ID & " : " & t.Name & " (Constraint: " & t.ConstraintType & ")" & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox constraintReport, vbInformation, "Constraints Report"
End Sub

```

### 3.3 Negative Float

Checks for tasks with TotalSlack < 0 (i.e., behind schedule or impossible logic).

```vb

Sub CheckNegativeFloat()
    Dim t As Task
    Dim negFloatReport As String
    negFloatReport = "Tasks with negative float:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.TotalSlack < 0 Then
                    negFloatReport = negFloatReport & _
                        " - " & t.ID & " : " & t.Name & " (Total Slack: " & t.TotalSlack & ")" & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox negFloatReport, vbInformation, "Negative Float Report"
End Sub

```

### 3.4 Tasks Without Resources

Finds non-summary tasks lacking resource assignments.

```vb

Sub CheckTasksWithoutResources()
    Dim t As Task
    Dim noResourceReport As String
    noResourceReport = "Tasks without resources:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.Assignments.Count = 0 Then
                    noResourceReport = noResourceReport & " - " & t.ID & " : " & t.Name & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox noResourceReport, vbInformation, "No Resource Report"
End Sub

```

### 3.5 Milestones with Durations

Flags tasks marked as milestones (t.Milestone = True) that have a non-zero duration.

```vb

Sub CheckMilestonesWithDuration()
    Dim t As Task
    Dim milestoneReport As String
    milestoneReport = "Milestones with non-zero duration:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.Milestone = True And t.Duration > 0 Then
                    milestoneReport = milestoneReport & _
                        " - " & t.ID & " : " & t.Name & " (Duration: " & t.Duration & " min)" & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox milestoneReport, vbInformation, "Milestone Duration Report"
End Sub

```

### 3.6 Summary Tasks with Resources

Reports summary tasks that have resource assignments (generally not recommended).

```vb

Sub CheckSummaryTasksWithResources()
    Dim t As Task
    Dim summaryResReport As String
    summaryResReport = "Summary tasks with resources assigned:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = True Then
                If t.Assignments.Count > 0 Then
                    summaryResReport = summaryResReport & _
                        " - " & t.ID & " : " & t.Name & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox summaryResReport, vbInformation, "Summary with Resources Report"
End Sub

```

### 3.7 Long-Duration Tasks

Flags tasks exceeding a certain duration (e.g., 20 working days).

```vb


Sub CheckLongDurationTasks()
    Dim t As Task
    Dim longDurationReport As String
    Dim DurationThreshold As Long

    ' Example: 20 working days = 20 * 480 = 9600 minutes
    DurationThreshold = 9600
    longDurationReport = "Tasks exceeding 20 working days:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.Duration > DurationThreshold Then
                    longDurationReport = longDurationReport & _
                        " - " & t.ID & " : " & t.Name & _
                        " (Duration: " & t.Duration / 480 & " days)" & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox longDurationReport, vbInformation, "Long Duration Report"

End Sub
```

### 3.8 Manually Scheduled Tasks

Highlights tasks set to Manually Scheduled instead of Auto Scheduled.

```vb

Sub CheckManuallyScheduledTasks()
    Dim t As Task
    Dim manualTasksReport As String
    manualTasksReport = "Manually Scheduled Tasks:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.Manual Then
                    manualTasksReport = manualTasksReport & _
                        " - " & t.ID & " : " & t.Name & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox manualTasksReport, vbInformation, "Manually Scheduled Tasks Report"
End Sub
```

### 3.9 Out-of-Sequence Progress

Checks if a task has actual progress but its predecessors are not completed (a rough indicator of out-of-sequence progress).

```vb

Sub CheckOutOfSequenceProgress()
    Dim t As Task
    Dim p As Task
    Dim oosReport As String
    oosReport = "Tasks with possible out-of-sequence progress:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.PercentComplete > 0 Then
                    ' Check each predecessor
                    For Each p In t.PredecessorTasks
                        If p.PercentComplete < 100 Then
                            oosReport = oosReport & _
                                " - " & t.ID & " : " & t.Name & _
                                " (Predecessor incomplete: " & p.ID & " : " & p.Name & ")" & vbCrLf
                            Exit For
                        End If
                    Next p
                End If
            End If
        End If
    Next t

    MsgBox oosReport, vbInformation, "Out-of-Sequence Progress Report"
End Sub
```

### 3.10 Missing Baseline

Finds tasks without Baseline Start or Baseline Finish (if a baseline is expected).

```vb

Sub CheckMissingBaseline()
    Dim t As Task
    Dim baselineReport As String
    baselineReport = "Tasks missing baseline data:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.BaselineStart = "NA" Or t.BaselineFinish = "NA" Then
                    baselineReport = baselineReport & _
                        " - " & t.ID & " : " & t.Name & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox baselineReport, vbInformation, "Missing Baseline Report"
End Sub

```

## 4. Additional / Advanced Checks

Below are macros for more nuanced checks, often referenced by DCMA 14-Point or GAO best practices:

- Excessive Leads/Lags
- Incorrect Dependency Types
- Over-/Under-Allocation
- Status Date Alignment (Actuals after the status date)
- Hard Constraints vs. Deadlines
- “Zombie” Tasks (100% complete but in the future)
- Resource Calendar Conflicts
- Critical Path Integrity
- Schedule Risk Factors (Low/No Slack)
- Data Date Progression

### 1) Excessive Leads or Lags

What it does:
Checks every task’s predecessors. If the lag is above a certain threshold or a negative lag (lead) is more negative than allowed, it flags them.

```vb
Sub CheckExcessiveLeadsOrLags()
    Dim t As Task
    Dim tr As TaskDependency
    Dim reportMsg As String
    Dim lagThreshold As Long

    ' For example, 2 days = 960 minutes
    lagThreshold = 960

    reportMsg = "Excessive Leads or Lags Report:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                ' Iterate over TaskDependencies (predecessor relationships)
                For Each tr In t.TaskDependencies
                    ' Positive Lag
                    If tr.Lag > lagThreshold Then
                        reportMsg = reportMsg & _
                            "- Task " & t.ID & " (" & t.Name & ") has a large LAG (" & tr.Lag / 480 & " days) " & _
                            "from predecessor " & tr.From.ID & " (" & tr.From.Name & ")" & vbCrLf
                    ' Negative Lag (Lead)
                    ElseIf tr.Lag < 0 Then
                        If tr.Lag < -480 Then
                            reportMsg = reportMsg & _
                                "- Task " & t.ID & " (" & t.Name & ") has a large LEAD (" & tr.Lag / 480 & " days) " & _
                                "from predecessor " & tr.From.ID & " (" & tr.From.Name & ")" & vbCrLf
                        End If
                    End If
                Next tr
            End If
        End If
    Next t

    MsgBox reportMsg, vbInformation, "Excessive Leads or Lags"
End Sub
```

### 2) Incorrect Dependency Types

What it does:
Flags tasks that use “SS” (Start-to-Start), “FF” (Finish-to-Finish), or “SF” (Start-to-Finish) dependency types.
Many organizations prefer Finish-to-Start as the default, unless there’s a justifiable reason for an alternative.

```vb
Sub CheckDependencyTypes()
    Dim t As Task
    Dim tr As TaskRelation
    Dim reportMsg As String

    reportMsg = "Non-Finish-to-Start Dependencies Report:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                For Each tr In t.TaskDependencies
                    Select Case tr.Type
                        Case pjStartToStart, pjFinishToFinish, pjStartToFinish
                            reportMsg = reportMsg & _
                                "- Task " & t.ID & " (" & t.Name & ") has a " & _
                                DependencyTypeName(tr.Type) & " dependency with " & _
                                tr.From.ID & " (" & tr.From.Name & ")" & vbCrLf
                    End Select
                Next tr
            End If
        End If
    Next t

    MsgBox reportMsg, vbInformation, "Dependency Type Report"
End Sub

Private Function DependencyTypeName(depType As PjTaskLinkType) As String
    Select Case depType
        Case pjFinishToFinish: DependencyTypeName = "Finish-to-Finish (FF)"
        Case pjStartToStart: DependencyTypeName = "Start-to-Start (SS)"
        Case pjStartToFinish: DependencyTypeName = "Start-to-Finish (SF)"
        Case Else: DependencyTypeName = "Finish-to-Start (FS)"
    End Select
End Function


```

### 3) Resources Over-Allocated or Under-Allocated

What it does:
Loops through each Resource in the project. Checks if they are overallocated according to MS Project’s logic.

Under-allocation is subjective; for example, you might define “under-allocated” if a resource is assigned less than 50% over the entire project timeline, or you may want to compare total availability vs. total work assigned.
Note: This a simple check for over-allocation. For more detailed analysis, you might compare `Resource.MaxUnits` vs. `Resource.Work`.

```vb
Sub CheckResourceOverallocation()
    Dim r As Resource
    Dim reportMsg As String

    reportMsg = "Over-Allocated Resources Report:" & vbCrLf

    For Each r In ActiveProject.Resources
        If Not r Is Nothing Then
            ' Overallocated property set by Project if assignment demands exceed resource capacity
            If r.OverAllocated = True Then
                reportMsg = reportMsg & _
                    "- Resource " & r.ID & " (" & r.Name & ") is Over-Allocated." & vbCrLf
            End If
        End If
    Next r

    MsgBox reportMsg, vbInformation, "Overallocation Check"
End Sub

' Example of a simple "under-allocation" check (very subjective)
Sub CheckResourceUnderallocation()
    Dim r As Resource
    Dim totalWork As Long
    Dim totalAvailability As Long
    Dim reportMsg As String

    reportMsg = "Under-Allocated Resources Report:" & vbCrLf

    For Each r In ActiveProject.Resources
        If Not r Is Nothing Then
            ' Work, Availability in minutes
            totalWork = r.Work
            totalAvailability = (r.MaxUnits / 100) * (ActiveProject.ProjectFinish - ActiveProject.ProjectStart) * 480
            ' For instance, flag if resource is allocated < 25% of their total availability
            If totalAvailability > 0 Then
                If (totalWork / totalAvailability) < 0.25 Then
                    reportMsg = reportMsg & _
                        "- Resource " & r.ID & " (" & r.Name & ") is Under-Allocated (<25%)." & vbCrLf
                End If
            End If
        End If
    Next r

    MsgBox reportMsg, vbInformation, "Underallocation Check"
End Sub
```

### 4) Incorrect Actual Dates (Status Date Alignment)

What it does:
Checks if any task has actual start/finish dates after the project’s status date (which is unrealistic if you’re trying to reflect “as of” that date). Also flags tasks that have an actual finish in the future relative to `Now()` (optional).

```vb

Sub CheckStatusDateAlignment()
    Dim t As Task
    Dim reportMsg As String
    Dim projStatusDate As Date

    ' If the project doesn't have a status date set, it might be "NA".
    If IsDate(ActiveProject.StatusDate) Then
        projStatusDate = ActiveProject.StatusDate
    Else
        ' If no StatusDate, fallback to today's date or exit
        projStatusDate = Date
    End If

    reportMsg = "Status Date Alignment Issues:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                ' Check actual start/finish
                If t.ActualStart <> "NA" Then
                    If t.ActualStart > projStatusDate Then
                        reportMsg = reportMsg & _
                            "- Task " & t.ID & " (" & t.Name & ") has Actual Start AFTER the Status Date." & vbCrLf
                    End If
                End If

                If t.ActualFinish <> "NA" Then
                    If t.ActualFinish > projStatusDate Then
                        reportMsg = reportMsg & _
                            "- Task " & t.ID & " (" & t.Name & ") has Actual Finish AFTER the Status Date." & vbCrLf
                    End If
                End If
            End If
        End If
    Next t

    MsgBox reportMsg, vbInformation, "Status Date Alignment Check"
End Sub
```

### 5) Deadlines vs. Constraints

What it does:
Flags tasks using hard constraints (e.g., Must Finish On) instead of recommended Deadlines.
A best practice is to set Deadline dates rather than using constraints that lock the schedule.

```vb
Sub CheckDeadlinesVsConstraints()
    Dim t As Task
    Dim reportMsg As String

    reportMsg = "Hard Constraints vs. Deadlines Report:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                ' Check if there's a "hard" constraint
                ' Hard constraints include Must Finish On (pjMFO), Must Start On (pjMSO), etc.
                Select Case t.ConstraintType
                    Case pjMFO, pjMSO, pjSNLT, pjFNLT, pjSNET, pjFNET
                        ' If there's also no Deadline, or a mismatch in dates
                        reportMsg = reportMsg & _
                            "- Task " & t.ID & " (" & t.Name & ") uses " & ConstraintTypeName(t.ConstraintType) & _
                            " constraint instead of a Deadline." & vbCrLf
                End Select
            End If
        End If
    Next t

    MsgBox reportMsg, vbInformation, "Hard Constraints vs. Deadlines"
End Sub

Private Function ConstraintTypeName(ct As Long) As String
    Select Case ct
        Case pjMFO: ConstraintTypeName = "Must Finish On (MFO)"
        Case pjMSO: ConstraintTypeName = "Must Start On (MSO)"
        Case pjSNLT: ConstraintTypeName = "Start No Later Than (SNLT)"
        Case pjFNLT: ConstraintTypeName = "Finish No Later Than (FNLT)"
        Case pjSNET: ConstraintTypeName = "Start No Earlier Than (SNET)"
        Case pjFNET: ConstraintTypeName = "Finish No Earlier Than (FNET)"
        Case Else: ConstraintTypeName = "As Soon As Possible or Other"
    End Select
End Function
```

### 6) “Zombie” Tasks (100% Complete but in the Future)

What it does:
Flags tasks that are marked 100% complete but whose finish date is in the future. This is a common data entry or status update mistake.

```vb
Sub CheckZombieTasks()
    Dim t As Task
    Dim reportMsg As String
    Dim todayDate As Date

    todayDate = Date  ' or use ActiveProject.StatusDate if you want to compare with status date

    reportMsg = "Zombie Tasks (100% Complete but Finish in Future):" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                If t.PercentComplete = 100 Then
                    ' If Finish is after today's date (or after the status date)
                    ' and the actual finish is not set or is after today
                    If t.Finish > todayDate Then
                        reportMsg = reportMsg & _
                            "- Task " & t.ID & " (" & t.Name & ") is 100% complete but finishes in the future." & vbCrLf
                    End If
                End If
            End If
        End If
    Next t

    MsgBox reportMsg, vbInformation, "Zombie Tasks Report"
End Sub
```

### 7) Resource Calendar Conflicts

What it does:
Flags tasks that overlap with a resource’s non-working time. In MS Project, if a resource is assigned to a task during that resource’s calendar “off” days, Project typically extends the task or shows overallocation.
Below is a basic example: check if a task’s scheduled working period conflicts with the resource’s calendar. A thorough check might need to iterate over each assignment’s time-phased data.

```vb
Sub CheckResourceCalendarConflicts()
    Dim t As Task
    Dim a As Assignment
    Dim reportMsg As String
    Dim assignmentWorkRange As String

    reportMsg = "Resource Calendar Conflicts:" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                For Each a In t.Assignments
                    If Not a.Resource Is Nothing Then
                        ' MS Project automatically adjusts for resource calendar,
                        ' so direct "conflict" is tricky to detect
                        '
                        ' We'll flag if the assignment spans days that are
                        ' recognized as non-working by the resource's calendar.

                        ' "CalendarConflict" is not a built-in property we can read directly,
                        ' so let's do a minimal check: if the resource is overallocated
                        ' specifically for this assignment or if work is scheduled
                        ' on a known non-working day, you'd need time-phased analysis.

                        If a.Resource.OverAllocated = True And a.Work > 0 Then
                            reportMsg = reportMsg & _
                                "- Task " & t.ID & " (" & t.Name & ") / Resource (" & a.Resource.Name & _
                                ") might have a calendar conflict or overallocation." & vbCrLf
                        End If
                    End If
                Next a
            End If
        End If
    Next t

    MsgBox reportMsg, vbInformation, "Resource Calendar Conflicts"
End Sub
```

**Deeper Approach:**

Iterate over each day in the task’s duration.
Check if that day is working time in the resource’s calendar.
If the schedule assigns work on a resource’s off day, flag it as a conflict.

### 8) Critical Path Integrity

What it does:
Lists tasks that are marked “critical” by MS Project. You can quickly inspect if the critical path makes sense or if there are unexpected tasks on the critical path.
You could also check if the project has no critical tasks (which would be odd if the project has a finish date).

```vb
Sub CheckCriticalPathIntegrity()
    Dim t As Task
    Dim reportMsg As String
    Dim criticalCount As Long

    reportMsg = "Critical Path Tasks:" & vbCrLf
    criticalCount = 0

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" And t.Critical = True Then
                criticalCount = criticalCount + 1
                reportMsg = reportMsg & _
                    "- Task " & t.ID & " (" & t.Name & ") is on the Critical Path." & vbCrLf
            End If
        End If
    Next t

    If criticalCount = 0 Then
        reportMsg = "No tasks marked as Critical. Is that expected?"
    End If

    MsgBox reportMsg, vbInformation, "Critical Path Integrity"
End Sub
```

### 9) Schedule Risk or Sensitivity Analysis

What it does:
A true risk/sensitivity (Monte Carlo) analysis typically requires add-in software (e.g., Barbecana Full Monte, Deltek Acumen). Basic MS Project VBA doesn’t directly simulate multiple scenarios.
However, we can show a simple macro that flags tasks with extremely low or zero total slack, indicating high sensitivity to delays.

```vb
Sub CheckScheduleRiskFactors()
    Dim t As Task
    Dim reportMsg As String

    reportMsg = "Potential High-Risk Tasks (Low/No Slack):" & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Summary = False And t.Name <> "" Then
                ' Slack is in minutes. Zero or near-zero slack is an indicator of risk.
                If t.TotalSlack <= 0 Then
                    reportMsg = reportMsg & _
                        "- Task " & t.ID & " (" & t.Name & ") has zero or negative Slack. (Slack = " & t.TotalSlack & ")" & vbCrLf
                ElseIf t.TotalSlack <= 240 Then
                    ' 240 min = half a day of slack; you can set your own threshold
                    reportMsg = reportMsg & _
                        "- Task " & t.ID & " (" & t.Name & ") has very low Slack (" & t.TotalSlack & " min)." & vbCrLf
                End If
            End If
        End If
    Next t

    MsgBox reportMsg, vbInformation, "Schedule Risk Factors"
End Sub
```

### 10) Data Date Progression

What it does:
Checks if the project’s status date is more than X days behind the current date, which might indicate the schedule isn’t being maintained regularly.

```vb
Sub CheckDataDateProgression()
    Dim projStatusDate As Date
    Dim daysBehind As Long
    Dim threshold As Long

    threshold = 7  ' e.g., if status date is more than a week behind

    If IsDate(ActiveProject.StatusDate) Then
        projStatusDate = ActiveProject.StatusDate
    Else
        MsgBox "Status Date is not set (NA).", vbInformation, "Data Date Progression Check"
        Exit Sub
    End If

    daysBehind = DateDiff("d", projStatusDate, Date)  ' Compare to today's date

    If daysBehind > threshold Then
        MsgBox "Project Status Date (" & Format(projStatusDate, "mm/dd/yyyy") & _
               ") is " & daysBehind & " days behind today’s date." & vbCrLf & _
               "Suggest updating the schedule status date.", vbExclamation, "Data Date Check"
    Else
        MsgBox "Project Status Date is up to date (within " & threshold & " days).", vbInformation, "Data Date Check"
    End If
End Sub
```

**Bringing It All Together**
You now have separate macros for each extended check. In practice, you might:

Create a “Master” macro that calls all these checks in sequence:

```vb
Sub RunAllExtendedChecks()
    CheckExcessiveLeadsOrLags
    CheckDependencyTypes
    CheckResourceOverallocation
    CheckResourceUnderallocation
    CheckStatusDateAlignment
    CheckDeadlinesVsConstraints
    CheckZombieTasks
    CheckResourceCalendarConflicts
    CheckCriticalPathIntegrity
    CheckScheduleRiskFactors
    CheckDataDateProgression
End Sub
```

Customize thresholds (e.g., lead/lag values, partial Slack limits) to match your organizational policy.

Enhance the reporting by writing the results to an `Excel` file, a `Word` doc, or a custom form in `MS Project`, rather than using simple message boxes.

Run the checks regularly (weekly, bi-weekly, or after major schedule changes) to keep the schedule healthy and free of hidden issues.

### Final Thoughts

This suite of VBA macros provides a comprehensive extension to the standard checks (missing predecessors, constraints, etc.). By tailoring thresholds and combining the results into a single automated report, you can create a powerful Schedule Health Check tool that aligns with industry best practices—from PMI to DCMA to GAO.
