
Function rand_B(x as integer, y as integer) as integer
	Randomize
    rand_B = Int(y * Rnd) + x
End Function

Function predict_weather(clim as string, seas as string, loc as string, prev as string) as string
	
End Function


Sub Record_Day()
	dim rain as string
	dim temp as string
	dim winds as string
	dim windd as string
	
	rain = Range("D3").Value
	temp = Range("D4").Value
	winds = Range("D5").Value
	windd = Range("D6").Value
	
    Range("B3").Value = rain
    Range("B4").Value = temp
    Range("B5").Value = winds
    Range("B6").Value = windd
End Sub