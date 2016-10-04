Option Explicit

Function rand_B(lowerbound As Integer, upperbound As Integer) As Integer
	Randomize
	rand_B = Int((upperbound - lowerbound + 1) * Rnd()) + lowerbound
End Function

Function predict_weather(clim As String, seas As String, loc As String, prev As String) As String

End Function

Function temp(climate As Integer, location As Integer, season As Integer) as Integer
	Dim mult as Integer	'Season/climate temperature multiplier
	Dim elev as Integer ' Elevation
	Dim rand as Integer ' Randomness modifier
	'find base temperature
	Select Case climate
		case 1		' Arctic
			temp = 10
			mult = 4
		case 2		' Sub-Arctic
			temp = 30
			mult = 3
		case 3		' Temperate
			temp = 50
			mult = 2
		case 4  	' Dry 
			temp = 80
			mult = 2
		case 5 		' Tropical
			temp = 80
			mult = 1
		case else 	' make the color of A5 red
			Range("A5").Interior.ColorIndex = 3
	End Select
	'roll and apply season modifier
	mult = mult*(rand_B(1,4)+4)
	Select Case season
		case 1 		' Summer
			temp = temp + mult
		case 2 to 3 ' Spring/Fall
			temp = temp
		case 4		' Winter
			temp = temp - mult
		case else 	' make the color of A5 red
			Range("A5").Interior.ColorIndex = 3
	End Select
	'apply location modifier
	Select Case location
		case 1 		' Mountains
			'elev = InputBox("Enter Approximate Elevation (in 1000's of feet)")
			elev = Range("D5").Value/1000
			temp = temp - (4*elev)
		case 2 to 5,7
			temp = temp
		case 6		' Desert
			temp = temp + 10
		case 8
			if (season = 1) then temp = temp - 10
			if (season = 4) then temp = temp + 10
		case else 	' make the color of A5 red
			Range("A5").Interior.ColorIndex = 3
	End Select
	'apply randomness
	temp = temp + (rand_B(1,20)+rand_B(1,20)+rand_B(1,20)+rand_B(1,20)-42)
End Function

Function wind(location As Integer) as Integer
	'roll 300
	wind = rand_B(1,300)
	'determine weather
	wind = Application.WorksheetFunction.Vlookup(wind,Worksheets("Weather").Range("G2:K11"),2)
	'adjust for modifiers
	Select Case location
		case 1 	'  Mountains
			wind = wind + 1
		case 2 to 3, 6, 7
			wind = wind
		case 4		' Forest
			if (wind > 1) then wind = wind - 1
		case 5, 8 		' Oceans and Coast
			wind = rand_B(1,300)
			if ((wind > 8) and (wind < 212)) or (wind > 280) Then
				wind = Application.WorksheetFunction.Vlookup(wind,Worksheets("Weather").Range("G2:K11"),2)
				wind = wind + 1
			else
				wind = Application.WorksheetFunction.Vlookup(wind,Worksheets("Weather").Range("G2:K11"),2)
			End if
	End Select
End Function

Function rain(climate as integer, location as Integer, season as Integer) as Integer
	'roll d100
	rain = rand_B(1,100)
	'add modifiers (season, location, climate)
	Select Case climate
		case 1 to 3		'Arctic, Sub-Arctic, Temperate
			rain = rain
		case 4 			'Dry
			rain = rain - 20
		case 5 			'Tropical
			rain = rain + 10
	End Select
	Select Case location
		case 1 to 4
			rain = rain
		case 6		' Desert
			rain = rain - 20
		case 7 		' Swamp
			rain = rain + 10
		case 5, 8 	' Coast/Ocean
			rain = rain + 5
	End Select
	Select Case season
		case 1,3
			rain = rain
		case 2 		'Spring
			rain = rain + 10
		case 4 		'Winter
			rain = rain + 5
	End Select
End Function

Function storm(climate as Integer, location as Integer, temp as integer, rain as integer, wind as integer) as String
	Dim num as Integer		' used for randomness
	storm = ""
	' Determine fog
	if (rain > 56 and rain < 61) then  ' if overcast, look at wind level
		Select Case wind
			case 1			' no wind
				storm = "heavy fog"
			case 2 			' slight wind
				storm = "light fog"
			case 3 to 10	'wind -> no fog
		End Select
	End if
	' Determine Storms
	Select case Wind
		case 1 to 3
			storm = ""
		case 4 to 5
			if (rain > 60) then
				if (temp <= 32) then
					storm = "Blizzard"
				else
					if (rand_B(1,2) = 1) then storm = "Thunderstorm" else storm = "Mild Storm"
					num = rand_B(-135,65)
					if (num > temp) then storm = storm & ", Hail"
				End if
			End if
		case 6 to 7
			if (rain > 60) then
				if (temp <= 32) then
					storm = "Severe Blizzard"
				else
					storm = "Severe Thunderstorm"
					if (rand_B(-20,65) > temp) then storm = storm & ", Large Hail"
					if (rand_B(1,2) = 1) and (climate > 2)	then
						if (location = 5) or (location = 8) then 
							storm = storm & ", Weak or Distant Waterspout"
						else 
							storm = storm & ", Weak or Distant Tornado"
						End if
					End if
				End if
			else
				if (location = 6) or (climate = 4) then 
					storm = "Sandstorm"
				elseif (location < 3) and (rand_B(1,2) = 1) then
					storm = "Land/Mudslide Hazard"
				elseif (rain > 40) and (temp > 84) then
					storm = "Dry Thunderstorm"
				else
					storm = "Windstorm"
				End if
			End if
		case 8
			if (rain > 60) then
				if (temp <= 32) then
					storm = "Violent Blizzard"
				elseif (temp >= 85) and (rain < 91) then
					storm = "Dry Thunderstorm, Wildfire Hazard"
				else
					storm = "Super Thunderstorm"
					if (rand_B(0,65) > temp) then storm = storm & ", Large Hail"
					if (rand_B(1,2) = 1) and (climate > 2)	then
						if (location = 5) or (location = 8) then 
							storm = storm & ", Waterspout"
						else 
							storm = storm & ", Tornado"
						End if
					End if
				End if
			else
				if (location = 6) or (climate = 4) then 
					storm = "Sandstorm"
				elseif (rain > 40) and (temp >= 70) then
					storm = "Dry Thunderstorm, Wildfire Hazard"
				else
					storm = "Windstorm"
				End if
			End if
			if (rand_B(1,100) > 80) then	'seismic activity
				if (location <= 2) then 
					storm = storm & ", Rockslide/Avalanche Hazard"
					if (location = 1) and (rand_B(1,2) = 1) then storm = storm & ", Potential Volcanic Activity"
				elseif ((location = 8) or (location = 5)) then
					storm = storm & ", Tsunami"
				else
					storm = storm & ", Earthquake"
					if (rand_B(1,2) = 1) then storm = storm & ", Sinkhole Hazard"
				End if
			End if
		case 9,10
			if (rain > 60) then
				if (temp <= 32) then
					storm = "Violent Blizzard"
				elseif (temp >= 85) and (rain < 91) then
					storm = "Dry Thunderstorm, Wildfire Hazard"
				else
					storm = "Super Thunderstorm"
					if (rand_B(0,65) > temp) then storm = storm & ", Large Hail"
					if (climate > 1) then
						if (location = 5) or (location = 8) then 
							storm = storm & ", Waterspout"
						else 
							storm = storm & ", Tornado"
						End if
					End if
				End if
			else
				if (location = 6) or (climate = 4) then 
					storm = "Sandstorm"
				elseif (rain > 40) and (temp >= 70) then
					storm = "Dry Thunderstorm, Wildfire Hazard"
				else
					storm = "Windstorm, Downburst, Rogue Tornado/Waterspout"
				End if
			End if
			if (rand_B(1,100) > 50) then	'seismic activity
				if (location <= 2) then 
					storm = storm & ", Rockslide/Avalanche"
					if (location = 1) and ((10 - wind + rand_B(1,2)) < 3) then storm = storm & ", Potential Volcanic Activity"
				elseif ((location = 8) or (location = 5)) then
					storm = storm & ", Tsunami"
				else
					storm = storm
				End if
				storm = storm & ", Earthquake, Sinkhole Hazard"
			End if
	End Select
	' add flavor for rain and snow
	if (storm = "") and (rain > 60) then
		if (temp <= 32) then
			if (rain > 90) then storm = "Heavy Snow" else storm = "Light Snow"
		else
			if (rain > 90) then 
				storm = "Heavy Rain" 
			elseif (rain > 76) then 
				storm = "Light Rain" 
			else storm = "Drizzle"
			End If
		End if
	End if
	' temp as a value, rain as a d100 roll, wind as a d10 category
	'return a string
End Function

Sub Roll_Weather()
	Dim climate as String
	Dim clim as Integer
	Dim season as String
	Dim seas as Integer
	Dim location as String
	Dim loc as Integer
	Dim temperature as Integer
	Dim windT as Integer
	Dim rainT as Integer
	'Convert string values to integers
	climate = Range("B3").Value
	Select Case climate
		case "Arctic"
			clim = 1
		case "Sub-Arctic"
			clim = 2
		case "Temperate"
			clim = 3
		case "Dry"
			clim = 4
		case "Tropical"
			clim = 5
	End Select
	season = Range("C3").Value
	Select Case season
		case "Summer"
			seas = 1
		case "Spring"
			seas = 2
		case "Fall"
			seas = 3
		case "Winter"
			seas = 4
	End Select
	location = Range("D3").Value
	Select Case location
		case "Mountain"
			loc = 1
		case "Hills"
			loc = 2
		case "Plains"
			loc = 3
		case "Forest"
			loc = 4
		case "Ocean"
			loc = 5
		case "Desert"
			loc = 6
		case "Swamp"
			loc = 7
		case "Coast"
			loc = 8
	End Select
	temperature = temp(clim,loc,seas)  ' get temperature
	Range("I8").Value = temperature
	Range("I9").Value = Application.WorksheetFunction.Vlookup(temperature,Worksheets("Weather").Range("M2:N9"),2)
	rainT = rain(clim,loc,seas)
	Range("I10").Value = Application.WorksheetFunction.Vlookup(rainT,Worksheets("Weather").Range("O2:P7"),2)
	windT = wind(loc)
	Range("I11").Value = Application.WorksheetFunction.Vlookup(windT,Worksheets("Weather").Range("H2:K11"),2)
	Range("I12").Value = Application.WorksheetFunction.Vlookup(rand_B(1,8),Worksheets("Weather").Range("H2:L9"),5)
	Range("I13").Value = storm(clim,loc,temperature,rainT,windT)
End Sub

Sub Next_Day()
	Dim i as Integer
	Dim x(1 to 7) as String
	Dim list(1 To 7) As String
	list(1) = "C"
	list(2) = "D"
	list(3) = "E"
	list(4) = "F"
	list(5) = "G"
	list(6) = "H"
	list(7) = "I"
	
	For i = 1 To 6
		x(1) = Range(list(i + 1) & "7").Value
		x(2) = Range(list(i + 1) & "8").Value
		x(3) = Range(list(i + 1) & "9").Value
		x(4) = Range(list(i + 1) & "10").Value
		x(5) = Range(list(i + 1) & "11").Value
		x(6) = Range(list(i + 1) & "12").Value
		x(7) = Range(list(i + 1) & "13").Value
		
		
		Range(list(i) & "7").Value = x(1)
		Range(list(i) & "8").Value = x(2)
		Range(list(i) & "9").Value = x(3)
		Range(list(i) & "10").Value = x(4)
		Range(list(i) & "11").Value = x(5)
		Range(list(i) & "12").Value = x(6)
		Range(list(i) & "13").Value = x(7)
		Next
End Sub

Sub Clear_Week()
	Range("C7:I13").Value = ""
End Sub	

Sub Roll_Week()
	Dim i as Integer
	For i = 1 to 6
	Roll_Weather
	Next_Day
	Next
	Roll_Weather
End Sub
