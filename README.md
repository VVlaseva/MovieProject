# Movie Dashboard
## Table of Content
[Problem Statement](#problem-statement)
[Data Source](#data-source)
[Tools](#tools)
[Dashboard](#dashboard)
[M-Code](m-code)
[Recommendations](Recommendations)
### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres.
Your task is to build an Excel dashboard that provides insights into this dataset. The dashboard should help identify:
- The best-performing actors
- The top movies based on box office metrics
- Director performance
- Genre trends
- Seasonal patterns in movie performance
- Any additional insights that can guide future production decisions

The final dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.
### Data Source
Movie Data : The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about movie's performance (box office and budget), actors, directors and genres.
You can download the original datasource here: [Movies Dataset Exsel file](https://github.com/user-attachments/files/34507236/Movies_Data_Homework.xlsx)
### Tools
1. Power Query - I used Power Query for Data Cleaning
2. Excel - I used Excel for Data Analysis
3. Pivot Tables - for Creating the Dashboard and Visualisation

Data Cleaning and Preparation
- Data loading and inspection
- Handling errors, missing values
- Data cleaning and formating. The excel file after the data cleaning & preparation process can be downloaded here - [Movies Dashboard](https://github.com/user-attachments/files/34507622/Movies.xlsx)
### Dashboard
<img width="1158" height="1010" alt="Dashboard" src="https://github.com/user-attachments/assets/7c636839-bca1-4cea-917f-6bbbc76f4a96" />

### M-Code
```
let
Source = Excel.Workbook(File.Contents("C:\Users\lenovo\OneDrive\Рабочий стол\Data Analitics\курс\дз\Power query\Movies_Data_Homework.xlsx"), null, true),
#"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
#"Promoted Headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
#"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
#"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
#"Merged Queries" = Table.NestedJoin(#"Removed Columns", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
#"Expanded Genres" = Table.ExpandTableColumn(#"Merged Queries", "Genres", {"Genre"}, {"Genres.Genre"}),
#"Renamed Columns" = Table.RenameColumns(#"Expanded Genres",{{"Genres.Genre", "Genres.First"}}),
#"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.First", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
#"Merged Queries1" = Table.NestedJoin(#"Reordered Columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
#"Expanded Genres1" = Table.ExpandTableColumn(#"Merged Queries1", "Genres", {"Genre"}, {"Genres.Genre"}),
#"Renamed Columns1" = Table.RenameColumns(#"Expanded Genres1",{{"Genres.Genre", "Genres.Second"}}),
#"Reordered Columns1" = Table.ReorderColumns(#"Renamed Columns1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.First", "Genre_Second_ID", "Genres.Second", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
#"Merged Queries2" = Table.NestedJoin(#"Reordered Columns1", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
#"Expanded Directors" = Table.ExpandTableColumn(#"Merged Queries2", "Directors", {"Director"}, {"Directors.Director"}),
#"Renamed Columns2" = Table.RenameColumns(#"Expanded Directors",{{"Directors.Director", "Directors.First"}}),
#"Reordered Columns2" = Table.ReorderColumns(#"Renamed Columns2",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.First", "Genre_Second_ID", "Genres.Second", "Director_First_ID", "Directors.First", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
#"Merged Queries3" = Table.NestedJoin(#"Reordered Columns2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
#"Expanded Actors" = Table.ExpandTableColumn(#"Merged Queries3", "Actors", {"Actor"}, {"Actors.Actor"}),
#"Renamed Columns3" = Table.RenameColumns(#"Expanded Actors",{{"Actors.Actor", "Actors.First"}}),
#"Reordered Columns3" = Table.ReorderColumns(#"Renamed Columns3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.First", "Genre_Second_ID", "Genres.Second", "Director_First_ID", "Directors.First", "Cast_First_ID", "Actors.First", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
#"Merged Queries4" = Table.NestedJoin(#"Reordered Columns3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
#"Expanded Actors1" = Table.ExpandTableColumn(#"Merged Queries4", "Actors", {"Actor"}, {"Actors.Actor"}),
#"Renamed Columns4" = Table.RenameColumns(#"Expanded Actors1",{{"Actors.Actor", "Actors.Second"}}),
#"Reordered Columns4" = Table.ReorderColumns(#"Renamed Columns4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.First", "Genre_Second_ID", "Genres.Second", "Director_First_ID", "Directors.First", "Cast_First_ID", "Actors.First", "Cast_Second_ID", "Actors.Second", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
#"Merged Queries5" = Table.NestedJoin(#"Reordered Columns4", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
#"Expanded Actors2" = Table.ExpandTableColumn(#"Merged Queries5", "Actors", {"Actor"}, {"Actors.Actor"}),
#"Renamed Columns5" = Table.RenameColumns(#"Expanded Actors2",{{"Actors.Actor", "Actors.Third"}}),
#"Reordered Columns5" = Table.ReorderColumns(#"Renamed Columns5",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.First", "Genre_Second_ID", "Genres.Second", "Director_First_ID", "Directors.First", "Cast_First_ID", "Actors.First", "Cast_Second_ID", "Actors.Second", "Cast_Third_ID", "Actors.Third", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
#"Merged Queries6" = Table.NestedJoin(#"Reordered Columns5", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
#"Expanded Actors3" = Table.ExpandTableColumn(#"Merged Queries6", "Actors", {"Actor"}, {"Actors.Actor"}),
#"Renamed Columns6" = Table.RenameColumns(#"Expanded Actors3",{{"Actors.Actor", "Actors.Fourth"}}),
#"Reordered Columns6" = Table.ReorderColumns(#"Renamed Columns6",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.First", "Genre_Second_ID", "Genres.Second", "Director_First_ID", "Directors.First", "Cast_First_ID", "Actors.First", "Cast_Second_ID", "Actors.Second", "Cast_Third_ID", "Actors.Third", "Cast_Fourth_ID", "Actors.Fourth", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
#"Merged Queries7" = Table.NestedJoin(#"Reordered Columns6", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
#"Expanded Actors4" = Table.ExpandTableColumn(#"Merged Queries7", "Actors", {"Actor"}, {"Actors.Actor"}),
#"Renamed Columns7" = Table.RenameColumns(#"Expanded Actors4",{{"Actors.Actor", "Actors.Fifth"}}),
#"Reordered Columns7" = Table.ReorderColumns(#"Renamed Columns7",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.First", "Genre_Second_ID", "Genres.Second", "Director_First_ID", "Directors.First", "Cast_First_ID", "Actors.First", "Cast_Second_ID", "Actors.Second", "Cast_Third_ID", "Actors.Third", "Cast_Fourth_ID", "Actors.Fourth", "Cast_Fifth_ID", "Actors.Fifth", "Budget ($)", "Box Office Revenue ($)"}),
#"Added Custom" = Table.AddColumn(#"Reordered Columns7", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
#"Filtered Rows" = Table.SelectRows(#"Added Custom", each true),
#"Changed Type1" = Table.TransformColumnTypes(#"Filtered Rows",{{"ROI", Percentage.Type}})
in
#"Changed Type1"
```

### Recommendations
<img width="360" height="146" alt="Снимок" src="https://github.com/user-attachments/assets/ac59add8-d33d-4d79-9fae-73abf69da7b6" />

