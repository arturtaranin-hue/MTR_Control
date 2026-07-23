VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaterialStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum MaterialReconciliationStatus
    PLANNED = 1
    PARTIAL = 2
    DELIVERED = 3
    OVERDELIVERED = 4
    NOT_FOUND = 5
    MANUAL_REVIEW = 6
End Enum

Public Function StatusName(ByVal status As MaterialReconciliationStatus) As String
    Select Case status
        Case PLANNED: StatusName = "PLANNED"
        Case PARTIAL: StatusName = "PARTIAL"
        Case DELIVERED: StatusName = "DELIVERED"
        Case OVERDELIVERED: StatusName = "OVERDELIVERED"
        Case NOT_FOUND: StatusName = "NOT_FOUND"
        Case MANUAL_REVIEW: StatusName = "MANUAL_REVIEW"
        Case Else: StatusName = "MANUAL_REVIEW"
    End Select
End Function
