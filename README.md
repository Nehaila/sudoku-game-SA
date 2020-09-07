## sudoku-game-SA
Sudoku game with simulated annealing using VBA

The code in VBA is as follows:

```
Private Sub SolveGame_Click()
T = 10000
imax = 1000
Tsfr = 100
proba = Exp(-((count1 - count0)) / (T + 0.1))
Random = Rnd
count0 = count()
count1 = count0
countmin = count()
i = 1
m = 1
'Start Simulated annealing
While T >= 0.001
For i = 1 To imax
    For m = 1 To Tsfr
        Sheets("sheet1").Range("K4").Value = i
        Sheets("sheet1").Range("K5").Value = m
        Sheets("sheet1").Range("K6").Value = countmin
        Sheets("sheet1").Range("K7").Value = count0
        Sheets("sheet1").Range("K8").Value = T
        
        x = Int(1 + Rnd * (9 - 1 + 1))
        y = Int(1 + Rnd * (9 - 1 + 1))

        Do Until Cells(x, y).Font.ColorIndex <> 37
            x = Int(1 + Rnd * (9 - 1 + 1))
            y = Int(1 + Rnd * (9 - 1 + 1))
        Loop

        xx = Int(1 + Rnd * (9 - 1 + 1))
        yy = Int(1 + Rnd * (9 - 1 + 1))

        Do Until Cells(xx, yy).Font.ColorIndex <> 37
            xx = Int(1 + Rnd * (9 - 1 + 1))
            yy = Int(1 + Rnd * (9 - 1 + 1))
        Loop

'Swapping the two random cells
    Swapp = Swap(x, y, xx, yy)
'count the new cost
    count1 = count()
'Test
    If countmin = 0 Then
        Exit Sub
    End If

    If count1 - count0 < 0 Then
        count0 = count1
        If count1 < countmin Then
            countmin = count1
        End If
        
    ElseIf count1 - count0 >= 0 Then
        If Random >= proba Then
            permute = Swap(x, y, xx, yy)
            count0 = count()
        ElseIf Random < proba Then
            count0 = count1
        End If
    End If
    Next
Next

T = T * 0.999
Wend
End Sub

Private Sub UserForm_Click()
End Sub



Function count()
ii = 1
jj = 1
p = 1
pp = 1
count = 0
count1 = 0
count2 = 0
countest = 0

'Count cost in columns
For ii = 1 To 9
p = 1
While p <= 9
countest = 0
    For jj = 1 To 9
            If Cells(ii, jj) = p Then
            countest = countest + 1
            End If
            
    Next
            If countest > 1 Then
                count1 = 1 + count1
            End If
    p = p + 1

Wend
    Next

'Count cost in Rows
jj = 1
ii = 1

For jj = 1 To 9
pp = 1
While pp <= 9
    countest = 0
    For ii = 1 To 9

                 If Cells(ii, jj) = pp Then
                    countest = countest + 1
                    
                End If
                
            
           
    Next
    If countest > 1 Then
            count2 = 1 + count2
            End If
    pp = pp + 1
Wend
Next
 count = count1 + count2
 countest = 0
 
 
 
 
 
 'If ii <> p Then
                 '       If Cells(ii, jj) = Cells(p, jj) Then
                  '      count1 = count1 + 1
                   '     End If
                'End If
                'p = p + 1
End Function

Function Erase2()
For f = 1 To 9
For m = 1 To 9
If Cells(f, m).Font.ColorIndex <> 37 Then
Cells(f, m).Clear
End If
Next
Next
End Function

Function Count3(iii, jjj)
count4 = 0
count5 = 0
count6 = 0
count7 = 0
count8 = 0
count9 = 0
count10 = 0
count11 = 0
count12 = 0
count13 = 0

'1ere Case

        If jjj <= 3 And 1 <= jjj Then
            If iii <= 3 And 1 <= iii Then
            
                For T = 1 To 3
                    For b = 1 To 3
                    If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count4 = count4 + 1
                        End If
                       
                     End If
                    Next
                Next
                End If
                End If
           
'2e Case
        If jjj <= 6 And 4 <= jjj Then
        If iii <= 3 And 1 <= iii Then
        
            T = 1
            b = 4
            
            For T = 1 To 3
                For b = 4 To 6
                    If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count5 = count5 + 1
                        End If
                        
                   End If
                Next
            Next
            End If
            End If
            
        
'3e Case
        If jjj <= 9 And 7 <= jjj Then
        If iii <= 3 And 1 <= iii Then
            T = 1
            b = 7
            
            For T = 1 To 3
                For b = 7 To 9
                If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count6 = count6 + 1
                        End If
                       
                        End If
                Next
            Next
            End If
            End If
            
        
'4e Case
        If jjj <= 3 And 1 <= jjj Then
        If iii <= 6 And 4 <= iii Then
            T = 4
            b = 1
            
            For T = 4 To 6
                For b = 1 To 3
                    If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count7 = count7 + 1
                        End If
                      
                    End If
                Next
            Next
            End If
            End If
            
       
'5e Case
        If jjj <= 6 And 4 <= jjj Then
        If iii <= 6 And 4 <= iii Then
            T = 4
            b = 4
            
            For T = 4 To 6
                For b = 4 To 6
                    If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count8 = count8 + 1
                        End If
                        
                    End If
                Next
            Next
            End If
            End If
            
        
'6eCase
        If jjj <= 9 And 7 <= jjj Then
        If iii <= 6 And 4 <= iii Then
            T = 4
            b = 7
            
            For T = 4 To 6
                For b = 7 To 9
                    If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count9 = count9 + 1
                        End If
                    
                    End If
                Next
            Next
            End If
            End If
        
        
'7e Case
        
If jjj <= 3 And 1 <= jjj Then
If iii <= 9 And 7 <= iii Then
            T = 7
            b = 1
            
            For T = 7 To 9
                For b = 1 To 3
                    If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count10 = count10 + 1
                        End If
                      
                  End If
                Next
            Next
            End If
            End If
        
'8e Case
If jjj <= 6 And 4 <= jjj Then
If iii <= 9 And 7 <= iii Then
            T = 7
            b = 4
            
            For T = 7 To 9
                For b = 4 To 6
                    If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count11 = count11 + 1
                        End If
                       
                   End If
                Next
            Next
            End If
            End If
        
'9e Case
If jjj <= 9 And 7 <= jjj Then
If iii <= 9 And 7 <= iii Then
            T = 7
            b = 7
            
            For T = 7 To 9
                For b = 7 To 9
                    If T <> iii Or b <> jjj Then
                        If Cells(T, b) = Cells(iii, jjj) Then
                        count13 = count13 + 1
                        End If
                       
                   End If
                Next
            Next
            End If
            End If
        
    

count12 = count4 + count5 + count6 + count7 + count8 + count9 + count10 + count11 + count13

Count3 = count12





End Function



Function Step1()
For r = 1 To 9
For T = 1 To 9
Randomize

If IsEmpty(Cells(r, T)) = True Then
rr = Int(1 + Rnd * (9 - 1 + 1))
Cells(r, T) = rr

Do Until Count3(r, T) = 0
rr = Int(1 + Rnd * (9 - 1 + 1))
Cells(r, T) = rr
Loop

End If
Next
Next


End Function

Function Swap(i, j, k, l)

Cells(400, 400) = Cells(i, j)
Cells(401, 401) = Cells(k, l)
Cells(k, l) = Cells(400, 400)
Cells(i, j) = Cells(401, 401)

End Function
```
You can find the entire code by checking the app. 
However, I would say that this is not the best way to solve a sudoku problem because it takes a big amount of time to get to the solution, given that we have many permutations to make and the probability of finding the exact one is usually low. Constraint programming would be better at solving these kind of problems. 