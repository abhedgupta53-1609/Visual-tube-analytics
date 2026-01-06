# Visual-tube-analytics
This is a data analysis project which give analysis of visual tube an imaginary video uploading platform and also visualize it.
So this project contains an excel which is macro enabled and works just like the ui of an analytical app.
In this Excel file there are 3 Dashboards 1st visualize the demographic data of cretors on the video uploading platform visual tube. 
2nd Dashboard Visualize the video meta data which makes it easy to know which type of video content was uploaded the most, which genere has most viewership,etc.
3rd Dashboard visualize the data of revenue generation like which genere generated the most revenue, how much revenue did creators earned from the platform, how much revenue did visual tube made and which age group of creators made the most revenue.

Here are some screenshots:-

<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_50_44 PM" src="https://github.com/user-attachments/assets/f0205217-fd55-49b4-b302-c27d7e8bb9dc" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_50_36 PM" src="https://github.com/user-attachments/assets/5d1d224c-181a-43e2-9553-813d1df13d9b" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_50_55 PM" src="https://github.com/user-attachments/assets/1d68a39e-652a-4867-8f3a-5336665cefbb" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_51_00 PM" src="https://github.com/user-attachments/assets/d8f7a723-62ae-4752-8d91-89648c354f83" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_54_33 PM" src="https://github.com/user-attachments/assets/e826f811-2ea5-4987-8243-4db785936ff0" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_54_13 PM" src="https://github.com/user-attachments/assets/eeebe33f-cd49-40e1-8e19-afebf142c8ad" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_54_03 PM" src="https://github.com/user-attachments/assets/82d8adf6-b41c-4481-8d4b-4d4b71a5c962" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_53_35 PM" src="https://github.com/user-attachments/assets/f5013699-19d0-4f63-992f-b70597d6691f" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_53_19 PM" src="https://github.com/user-attachments/assets/0e9d30f7-0e61-4915-a711-0583ca135755" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_53_05 PM" src="https://github.com/user-attachments/assets/bb188e13-5ae7-49bd-89f9-2650aa55d217" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_52_55 PM" src="https://github.com/user-attachments/assets/ee750923-caa2-48f3-bc4e-db5d0f2ed1fa" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_52_47 PM" src="https://github.com/user-attachments/assets/c683c418-5b40-4032-8430-6059067cac25" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_52_38 PM" src="https://github.com/user-attachments/assets/038edd29-9bf3-4409-9dbe-8a5bdf13c9ec" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_52_33 PM" src="https://github.com/user-attachments/assets/f3a9c3cf-e1a6-4aa4-ad63-a6c77972e822" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_52_18 PM" src="https://github.com/user-attachments/assets/bd39ab2c-ed2f-42ea-982a-ecaad9142c99" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_52_12 PM" src="https://github.com/user-attachments/assets/c47ec49b-e90b-44ec-9739-c7f49210fb55" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_51_58 PM" src="https://github.com/user-attachments/assets/9ae26ca3-27cc-49d9-b1bb-2252fb47d62f" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_51_49 PM" src="https://github.com/user-attachments/assets/d358b1e6-a47e-4fae-a1a6-c25a80b8a346" />
<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_51_18 PM" src="https://github.com/user-attachments/assets/f9971312-2d70-48ae-afd5-c62777be4220" />

<img width="1920" height="1020" alt="Visual tube macro - Excel 1_1_2026 4_50_48 PM" src="https://github.com/user-attachments/assets/bca50712-013f-4644-afa5-358e7a93ade9" />

macros used in this project:-
-- Macros from 1st Dashboard

Sub Show_Slicer_1()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Slicer")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide")).Visible = msoTrue
End Sub

Sub Unshow_Slicer_1()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Slicer")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide")).Visible = msoTrue
End Sub

Sub Show_Country_Creator()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Country & Creator")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 1")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 1")).Visible = msoTrue
End Sub

Sub Unshow_Country_Creator()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Country & Creator")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 1")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 1")).Visible = msoFalse
    
End Sub
Sub Show_Age_Group_Creator()
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Age Group & Creator")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 2")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 2")).Visible = msoFalse
    
End Sub

Sub Unshow_Age_Group_Creator()
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Age Group & Creator")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 2")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 2")).Visible = msoTrue
    
End Sub

Sub Show_Agegroup_Gender_Creator()
'
' Macro3 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Age group, Gender & Creators")).Visible = _
        msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 3")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 3")).Visible = msoTrue
End Sub

Sub Unshow_Agegroup_Gender_Creator()
'
' Macro3 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Age group, Gender & Creators")).Visible = _
        msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 3")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 3")).Visible = msoFalse
    
End Sub

-- Macros from 2nd Dashboard

Sub Show_dates()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Slicer")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide")).Visible = msoTrue
End Sub

Sub Unshow_dates()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Slicer")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide")).Visible = msoFalse
    
End Sub
Sub Show_country_uploads()
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Country & No. of Uploads")).Visible = _
        msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 1")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 1")).Visible = msoTrue
End Sub

Sub Unshow_country_uploads()
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Country & No. of Uploads")).Visible = _
        msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 1")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 1")).Visible = msoFalse
    
End Sub
Sub Show_vid_category_uploads()
'
' Macro3 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Video Category & No. of Uploads")).Visible _
        = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 2")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 2")).Visible = msoTrue
End Sub

Sub Unshow_vid_category_uploads()
'
' Macro3 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Video Category & No. of Uploads")).Visible _
        = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 2")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 2")).Visible = msoFalse
    
End Sub


Sub Show_category_watctime()
'
' Macro4 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Category & Watchtime")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 3")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 3")).Visible = msoTrue
End Sub
Sub Unshow_category_watctime()
'
' Macro4 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Category & Watchtime")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 3")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 3")).Visible = msoFalse
    
End Sub
Sub Show_view_gender()
'
' Macro5 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Views & Gender")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 4")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 4")).Visible = msoTrue
End Sub
Sub Unshow_view_gender()
'
' Macro5 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Views & Gender")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 4")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 4")).Visible = msoFalse
    
End Sub
Sub Show_Avg_watctime_gender()
'
' Macro6 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Avg. Watchtime by Gender")).Visible = _
        msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 5")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 5")).Visible = msoTrue
End Sub
Sub Unshow_Avg_watctime_gender()
'
' Macro6 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Avg. Watchtime by Gender")).Visible = _
        msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 5")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 5")).Visible = msoFalse
    
End Sub
Sub Show_number_upload_length()
'
' Macro7 Macro
'

'
    ActiveSheet.Shapes.Range(Array("No. of Uploads & Video length")).Visible = _
        msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 6")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Unhide 6")).Visible = msoTrue
End Sub
Sub Unshow_number_upload_length()
'
' Macro7 Macro
'

'
    ActiveSheet.Shapes.Range(Array("No. of Uploads & Video length")).Visible = _
        msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 6")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Unhide 6")).Visible = msoFalse
    
End Sub

-- Macros from 3rd Dashboard

Sub Show_gender_revenue()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Gender & Revenue")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("UnHide")).Visible = msoTrue
End Sub
Sub Unshow_gender_revenue()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Gender & Revenue")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("UnHide")).Visible = msoFalse
    
End Sub

Sub Show_vidlen_revenue()
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Revenue By Video Length")).Visible = _
        msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 1")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("UnHide 1")).Visible = msoTrue
End Sub
Sub Unshow_vidlen_revenue()
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Revenue By Video Length")).Visible = _
        msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 1")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("UnHide 1")).Visible = msoFalse
    
End Sub

Sub Show_vidcateg_revenue()
'
' Macro3 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Video Category & Revenue")).Visible = _
        msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 2")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("UnHide 2")).Visible = msoTrue
End Sub
Sub Unshow_vidcateg_revenue()
'
' Macro3 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Video Category & Revenue")).Visible = _
        msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 2")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("UnHide 2")).Visible = msoFalse
    
End Sub
Sub Show_scatterplot_revenue()
'
' Macro4 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Scatter Plot Group")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 3")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("UnHide 3")).Visible = msoTrue
End Sub
Sub Unshow_scatterplot_revenue()
'
' Macro4 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Scatter Plot Group")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 3")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("UnHide 3")).Visible = msoFalse
    
End Sub
Sub Show_slicergrp_date()
'
' Macro5 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Slicer Group Dates")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Hide 4")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("UnHide 4")).Visible = msoTrue
End Sub
Sub Unshow_slicergrp_date()
'
' Macro5 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Slicer Group Dates")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Hide 4")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("UnHide 4")).Visible = msoFalse
    
End Sub

Thanks for Reading :-)
