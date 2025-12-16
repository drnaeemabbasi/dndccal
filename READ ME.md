DNDCv.CAN Calibration System
Step-by-Step User Guide
BEFORE YOU START
You need:
•	DNDC95.exe installed at C:\DNDC
•	A working .dnd file that runs successfully
•	A batch file for DNDC
•	Observed data (yield, NEE, ET, or soil measurements)
STEP 1: PREPARE YOUR OBSERVED DATA FILE
Option A: Download the template
1.	Open the calibration tool
2.	Select your target variable (Yield, SoilTemp, SoilMoisture, ET, or NEE)
3.	Click "Obs. Template" and save the CSV template
Option B: Create it yourself
For Yield:
Year,Value
2015,4500
2016,4800
2017,4200	

For daily data (NEE, ET, soil temp/moisture):
Year,Day,Value
2015,120,5.2
2015,121,6.1
2015,122,5.8
Rules:
•	First two rows are headers (skip rows in template = 2)
•	No missing values - remove incomplete rows
•	Dates must match your DNDC simulation period
STEP 2: PREPARE YOUR PARAMETER FILE
Download the template:
1.	Click "Param. Template"
2.	Fill in your parameters
Format:
parameter_name,min,max,line_number
crop_growth_rate,0.8,1.5,245
soil_respiration,0.5,2.0,178
nitrogen_uptake,0.6,1.4,312
How to find line_number:
1.	Open your .dnd file in Notepad
2.	Find the parameter you want to calibrate
3.	Count the line number (starts at 1)
4.	The value you want to change should be in position 2 on that line
Tips:
•	Start with 2-4 parameters max
•	Use realistic bounds from literature
•	Too wide = slow convergence
•	Too narrow = might miss optimal value
STEP 3: RUN THE CALIBRATION
1.	Open the calibration tool
2.	Select Target Variable:
Choose: Yield, SoilTemp, SoilMoisture, ET, or NEE
If soil variable: select depth (e.g., 10cm, 20cm)
3.	Load Files:
Batch File: Browse to your .txt batch file
.dnd File: Browse to your .dnd file
Observed CSV: Browse to your observed data file
Parameter CSV: Browse to your parameter bounds file
4.	Set Iterations:
Default: 10 (quick test)
Recommended: 50-100 (proper calibration)
More iterations = better results but slower
5.	Click "Start Calibration"
6.	Monitor Progress:
Watch the log window
Green progress bar shows % complete
Each iteration shows: parameters tested, RMSE achieved
Best RMSE is highlighted when found
7.	Wait or Stop:
Let it run to completion (recommended)
Or click "Stop" anytime (results saved up to that point)
STEP 4: CHECK YOUR RESULTS
File location: C:\DNDC\[target_variable]_calibration_results.xlsx
Examples:
•	yield_calibration_results.xlsx
•	soil_temp_10cm_results.xlsx
•	nee_calibration_results.xlsx
Open the Excel file - three sheets:
Sheet 1: "All Iterations"
•	Every parameter combination tested
•	RMSE, R², MAE for each
•	Gold row = best result
Sheet 2: "Best Iteration"
•	Just the optimal parameters
•	Copy these values to your .dnd file for production runs
Sheet 3: "Data Comparison"
•	Observed vs. Modeled values
•	Chart (for daily data)
•	Check fit visually
STEP 5: USE YOUR CALIBRATED PARAMETERS
Option A: Manual update
1.	Open your .dnd file
2.	Go to each line_number from parameter file
3.	Replace old values with values from "Best Iteration" sheet
4.	Save and run DNDC normally
Option B: Use backup files
1.	Go to C:\DNDC\dnd_backups\
2.	Find the iteration you want (e.g., iter_25.dnd if iteration 25 was best)
3.	Copy it over your original .dnd file
4.	Run DNDC
Additional notes (recommendations)
•	Start small: 2-3 parameters, 20 iterations, test first
•	Check manually: Run DNDC with mid-range parameters before calibrating
•	Use literature: Base your bounds on published values
•	Be patient: 100 iterations × 2 min/run = 3-4 hours
•	Save backups: Keep your original .dnd file safe
•	Verify results: Run DNDC manually with best parameters to confirm
QUICK START CHECKLIST
☐	DNDC installed and working
☐	.dnd file runs successfully
☐	Observed data prepared (correct format)
☐	Parameters identified
☐	Bounds researched (realistic ranges)
☐	Parameter CSV created (line numbers correct)
☐	Iterations set
