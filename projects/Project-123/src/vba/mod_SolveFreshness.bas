Attribute VB_Name = "mod_SolveFreshness"
Option Explicit

' Project 123 -- Solve freshness tracking
'
' StagedIDC and StagedDebtSize hold plain numbers, not formulas. Change an assumption
' without re-running the solve and every downstream figure still calculates happily off
' the old staged values -- the model looks right and is wrong, with nothing on screen to
' say so. That is the single most dangerous failure mode in a macro-solved model.
'
' RecordSolveSnapshot copies the live value of every tracked input into the snapshot
' column on Cover. The Match column and Solve Status then flag any drift, and name the
' assumption that moved rather than just reporting "something changed".
'
' Every solve macro calls this on success. It should not be called anywhere else --
' recording a snapshot without actually solving is exactly the lie this guards against.

Public Sub RecordSolveSnapshot()
    Range("SnapshotStored").Value = Range("SnapshotLive").Value
    Range("LastSolvedStamp").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Application.Calculate
End Sub


' Clears the snapshot so the model reports as unsolved. Use after restoring a backup or
' when you want to force a visible re-solve prompt.
Public Sub InvalidateSolveSnapshot()
    Range("SnapshotStored").ClearContents
    Range("LastSolvedStamp").Value = "(never)"
    Application.Calculate
    MsgBox "Solve snapshot cleared. The model will report as needing a re-solve.", _
           vbInformation, "Freshness -- Invalidated"
End Sub
