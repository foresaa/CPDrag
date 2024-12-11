# &copy; 2024 Andy Forrester, Andel Projects

## All rights reserved.

This VBA script calculates the Critical Path Drag for each task in an MS Project file.

## Use of this code is permitted for educational and professional purposes.

<img src="perpop.png" alt="Andel Projects Limited" width="200">

```vba

Sub CalculateCriticalPathDrag()
    ' Define the project object to work with the active project
    Dim proj As Project
    Set proj = ActiveProject

    ' Define task variables
    Dim t As Task
    Dim otherTask As Task
    Dim drag As Double
    Dim taskDuration As Double
    Dim parallelTasks As Collection
    Dim minParallelFloat As Double
    Dim dragReductionBenefit As Double
    Dim dayCost As Double

    ' Iterate through each task in the project
    For Each t In proj.Tasks
        ' Ensure the task is valid (not null)
        If Not t Is Nothing Then
            ' Process only tasks that are critical, are not milestones, and are not summary tasks
            If t.Critical And Not t.Milestone And Not t.Summary Then
                ' Step 1: Calculate the initial drag value for the task
                taskDuration = t.Duration / (60 * proj.HoursPerDay) ' Convert duration from minutes to days
                drag = taskDuration ' Start with the task duration as the initial drag value
                Set parallelTasks = New Collection ' Initialize collection to store parallel tasks

                ' Step 2: Identify parallel activities and calculate their total float
                Dim otherFloat As Double
                For Each otherTask In proj.Tasks
                    ' Ensure the other task is valid, not the same as the current task, not critical, not a milestone, and not a summary task
                    If Not otherTask Is Nothing And otherTask.UniqueID <> t.UniqueID And Not otherTask.Critical And Not otherTask.Milestone And Not otherTask.Summary Then
                        ' Check if the other task overlaps with the current task
                        If otherTask.Start < t.Finish And otherTask.Finish > t.Start Then
                            ' Calculate total float in days
                            otherFloat = otherTask.TotalSlack / (60 * proj.HoursPerDay) ' Convert slack from minutes to days
                            ' Add the other task to the collection of parallel tasks
                            parallelTasks.Add otherTask
                            ' Update the drag value if the total float of the parallel task is less than the current drag
                            If otherFloat < drag Then
                                drag = otherFloat
                            End If
                        End If
                    End If
                Next otherTask

                ' Step 3: Set the drag value in the custom field Number20
                t.Number20 = drag
                Debug.Print "Task: " & t.Name & ", Critical Path Drag: " & drag

                ' Step 4: Calculate the benefit of reducing the drag (Drag Reduction Benefit)
                If t.Duration > 0 Then
                    dayCost = t.Cost / (t.Duration / (60 * proj.HoursPerDay)) ' Calculate the cost per day of the task
                    dragReductionBenefit = dayCost * drag ' Calculate the drag reduction benefit
                    t.Cost1 = dragReductionBenefit ' Set Cost1 as the drag reduction benefit
                Else
                    dragReductionBenefit = 0
                End If

                ' Step 5: Print the minimum float for the parallel tasks, if any were found
                If parallelTasks.Count > 0 Then
                    Debug.Print "Minimum Total Float for Parallel Tasks of " & t.Name & ": " & drag
                End If
            End If
        End If
    Next t

    ' Display a message box to indicate that the process is complete
    MsgBox "Critical Path Drag calculation completed."
End Sub
```
