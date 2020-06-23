Attribute VB_Name = "Module1"

Sub clear()

  sheets("2014").Columns("I:R").EntireColumn.Delete
  sheets("2015").Columns("I:R").EntireColumn.Delete
  sheets("2016").Columns("I:R").EntireColumn.Delete

End Sub

