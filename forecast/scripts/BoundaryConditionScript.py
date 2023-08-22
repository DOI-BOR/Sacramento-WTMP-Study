import os, sys
import re
from com.rma.io import DssFileManagerImpl
from com.rma.model import Project

import hec.heclib.dss
import hec.heclib.util.HecTime as HecTime
import hec.io.TimeSeriesContainer as tscont
import hec.hecmath.TimeSeriesMath as tsmath

import usbr.wat.plugins.actionpanel.model.forecast as fc

import java.lang
import java.io.File
import java.io.FileInputStream
from org.apache.poi.xssf.usermodel import XSSFWorkbook
from org.apache.poi.hssf.usermodel import HSSFWorkbook
from org.apache.poi.ss import usermodel as SSUsermodel

sys.path.append(os.path.join(Project.getCurrentProject().getWorkspacePath(), "forecast", "scripts"))

import CVP_ops_tools as CVP
reload(CVP)

'''Accepts parameters for WTMP forecast runs to form boundary condition data sets.'''
def build_BC_data_sets(AP_start_time, AP_end_time, BC_F_part, BC_output_DSS_filename, ops_file_name, DSS_map_file,
		position_analysis_year=None,
		position_analysis_cofig_file=None,
		met_F_part=None,
		met_output_DSS_filename=None,
		flow_pattern_config_file=None,
		ops_import_F_part=None):

	# Postitional (required) args:
	# AP_start_time (HecTime) start of the simulation group run time
	# AP_end_time (HecTime) end of the simulation group run time
	# BC_F_part (str) DSS F part for output time series records
	# BC_output_DSS_filename (str) Name of DSS file for output time series records. Assumed relative to study directory
	# ops_file_name (str) Name of CVP ops data spreadsheet file
	# DSS_map_file (str) Name of file where list of output locactions and DSS records will be written.  Assumed relative to study directory

	# Key-word (optional) args (kwargs):
	# position_analysis_year (int) Source year for met data position analysis (positional analysis args are needed until there are other methods for making met data)
	# position_analysis_cofig_file (str) Name of file holding list of source time series for position analysis. Assumed relative to study directory. Defaults to forecast/config/historical_met.config
	# met_F_part (str) DSS F part for met data specifically. Defaults to BC_F_part
	# met_output_DSS_filename (str) Name of separate DSS file for met time series records. Assumed relative to study directory. Defaults to BC_output_DSS_filename
	# flow_pattern_config_file (str) Name of file holding list of pattern time series for flow disaggreagtion. Assumed relative to study directory. Defaults to forecast/config/flow_pattern.config


	if not os.path.isabs(BC_output_DSS_filename):
		BC_output_DSS_filename = os.path.join(Project.getCurrentProject().getWorkspacePath(), BC_output_DSS_filename)
	if not os.path.isabs(ops_file_name):
		ops_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), ops_file_name)
	if not met_F_part:
		met_F_part = BC_F_part
	if not ops_import_F_part:
		ops_import_F_part = BC_F_part
	if not met_output_DSS_filename:
		met_output_DSS_filename=BC_output_DSS_filename
	elif not os.path.isabs(met_output_DSS_filename):
		met_output_DSS_filename = os.path.join(Project.getCurrentProject().getWorkspacePath(), met_output_DSS_filename)
	if not position_analysis_cofig_file:
		position_analysis_cofig_file = fc.getHistoricalMetFile()
		#position_analysis_cofig_file = os.path.join(Project.getCurrentProject().getWorkspacePath(), "forecast", "config", "historical_met.config"
	elif not os.path.isabs(position_analysis_cofig_file):
		position_analysis_cofig_file = os.path.join(Project.getCurrentProject().getWorkspacePath(), position_analysis_cofig_file)
	# print "Historic Met Data config = " + position_analysis_cofig_file
	if not flow_pattern_config_file:
		flow_pattern_config_file = fc.getFlowPatternFile()
		# flow_pattern_config_file = os.path.join(Project.getCurrentProject().getWorkspacePath(), "forecast", "config", "flow_pattern.config")
	elif not os.path.isabs(flow_pattern_config_file):
		flow_pattern_config_file = os.path.join(Project.getCurrentProject().getWorkspacePath(), flow_pattern_config_file)
	# print "Flow Pattern Data config = " + flow_pattern_config_file
	if not os.path.isabs(DSS_map_file):
		DSS_map_file = os.path.join(Project.getCurrentProject().getWorkspacePath(), DSS_map_file)

	print "\n########"
	print "\tRunning Boundary Condition generation process"
	print "########\n"

	target_year = AP_end_time.year()

	met_lines = create_positional_analysis_met_data(target_year, position_analysis_year, AP_start_time, AP_end_time,
		position_analysis_cofig_file, met_output_DSS_filename, met_F_part)
	with open(os.path.join(Project.getCurrentProject().getWorkspacePath(), DSS_map_file), "w") as mapfile:
		mapfile.write("location,parameter,dss file,dss path\n")
		for line in met_lines:
			mapfile.write(line)
			mapfile.write('\n')

	ops_lines = create_ops_BC_data(target_year, ops_file_name, AP_start_time, AP_end_time,
		BC_output_DSS_filename, BC_F_part, ops_import_F_part, flow_pattern_config_file, DSS_map_file)
	with open(os.path.join(Project.getCurrentProject().getWorkspacePath(), DSS_map_file), "a") as mapfile:
		for line in ops_lines:
			mapfile.write(line)
			mapfile.write('\n')

	print "\nBoundary condition report written to: %s\n"%(DSS_map_file)

	return len(met_lines) + len(ops_lines)


'''Simple time-shifer for met positional ananlysis data'''
def create_positional_analysis_met_data(target_year, source_year, start_time, end_time,
position_analysis_cofig_file, met_output_DSS_filename, met_F_part):
	print "Calculating positional met data..."
	diff_years = target_year - source_year
	print "Shifting met data from %d to %d (%d years)."%(source_year, target_year, diff_years)

	rv_lines = []
	met_config_str = ""
	DSSout = hec.heclib.dss.HecDss.open(met_output_DSS_filename)
	with open(position_analysis_cofig_file) as met_config_file:
		met_config_str = met_config_file.read()
	met_config_str = re.sub(r"<!--.*?-->", '', met_config_str).strip()
	met_config_str = re.sub("\n\n", "\n", met_config_str)
	print met_config_str
	met_config_lines = met_config_str.split('\n')
	for line in met_config_lines:
		token = line.strip().split(',')
		if len(token) != 3:
			print "File %s line \n\t \"%s\"\nis not a valid ID for a position analysis DSS record."%(position_analysis_cofig_file,line)
			continue
		source_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), token[0].strip('\\'))
		DSSsource = hec.heclib.dss.HecDss.open(source_DSS_file_name)
		print "Reading %s from DSS file %s."%(token[2].strip(), source_DSS_file_name)
		tsmath_source = DSSsource.read(token[2].strip())
		time_step_label = token[2].strip().split('/')[5]
		# print "\tTime series contains %d values."%(tsmath_source.getContainer().numberValues)
		# print "\tShifting time series with shiftInTime(%s)."%("%dMo"%(diff_years*12))
		# tsmath_shift = tsmath_source.shiftInTime("%dYrar"%(diff_years))
		tsmath_shift = tsmath.generateRegularIntervalTimeSeries(
			"%s 0000"%(start_time.date(4)),
			"%s 2400"%(end_time.date(4)),
			time_step_label, "0M", 1.0)
		time_seek = HecTime(tsmath_shift.firstValidDate(), HecTime.MINUTE_INCREMENT)
		time_seek.setYearMonthDay(time_seek.year() - diff_years, time_seek.month(), time_seek.day(), time_seek.minutesSinceMidnight())
		if time_seek.getMinutes() < tsmath_source.firstValidDate():
			print "Met position time shift out of range at source start..."
			return ['']
		source_container = tsmath_source.getContainer()
		shift_container = tsmath_shift.getContainer()
		start_index = 0
		for i in range(source_container.numberValues):
			if source_container.times[i] >= time_seek.getMinutes():
				start_index = i
				break
		if start_index == 0:
			print "Met position time shift out of range at source end..."
			return ['']
		# if this works, it's only because the source and shift TSCs have the same time step.
		for i in range(shift_container.numberValues):
			shift_container.values[i] = source_container.values[start_index + i]
		if len(shift_container.values) != shift_container.numberValues:
			print "You doofus!\nlen(values)=%d\nnumberValues=%d\n"%(len(shift_container.values), shift_container.numberValues)
			return ['']
		tsmath_shift.setType(tsmath_source.getType())
		tsmath_shift.setUnits(tsmath_source.getUnits())
		tsmath_shift.setPathname(tsmath_source.getContainer().fullName)
		tsmath_shift.setVersion(met_F_part)
		DSSout.write(tsmath_shift)
		DSSsource.done()

		met_loc, met_param = token[1].strip().split('<', 1)
		rv_lines.append("%s,%s,%s,%s"%(met_loc.strip(), met_param.strip().strip('>'),
		Project.getCurrentProject().getRelativePath(met_output_DSS_filename),
		tsmath_shift.getContainer().fullName))

	return rv_lines

'''Processes the contents of the CVP ops spreadsheet in to flow and water temperature BCs'''
def create_ops_BC_data(target_year, ops_file_name, start_time, end_time, BC_output_DSS_filename, BC_F_part, ops_import_F_part, flow_pattern_config_file, DSS_map_file):
	print "Processing boundary conditions for Shasta Lake from ops file:\n\t%s"%(ops_file_name)

	forecast_locations = ["Trinity/Clair Engle", "Whiskeytown", "Shasta", "Oroville", "Folsom", "New Melones", " SAN LUIS/O'NEILL", "DELTA"]
	active_locations = ["Trinity/Clair Engle", "Whiskeytown", "Shasta"]

	rv_lines = []

	if ops_file_name.endswith(".xls") or ops_file_name.endswith(".xlsx"):
		ops_data = import_CVP_Ops_xls(ops_file_name, forecast_locations)
	else:
		ops_data = import_CVP_Ops_csv(ops_file_name, forecast_locations)
	# for key in ops_data.keys():
		# print "ops_data key: %s"%(key)

	shasta_ts_list = []
	shasta_calendar = ops_data["Shasta"][0].split(',')
	start_index = int(shasta_calendar[0])
	start_month = shasta_calendar[start_index + 1].strip().upper()
	# print "\n Shasta start month: %s; Start index: %d"%(start_month, start_index)
	for line in ops_data["Shasta"][1:]:
		# print "Passing line to CVP.make_ops_tsc: %s"%(line)
		data_month = start_month
		data_year = target_year
		try:
			early_val = float(line.split(',')[start_index - 1].strip())
			data_month = CVP.month_TLA[CVP.last_month(CVP.month_index(start_month))]
			if data_month == "DEC":
				data_year -= 1
		except:
			pass
		# print "Start_index = %d\nData_Month = %s"%(start_index, data_month)
		# print "Passing line to CVP.make_ops_tsc: %s"%(line)
		shasta_ts_list.append(CVP.make_ops_tsc("SHASTA", data_year, data_month, line, ops_label=ops_import_F_part))
		# ts_count += 1

	whiskeytown_ts_list = []
	whiskeytown_calendar = ops_data["Shasta"][0].split(',')
	whiskeytown_start_index = int(whiskeytown_calendar[0])
	whiskeytown_start_month = whiskeytown_calendar[start_index + 1].strip().upper()
	# print "\n Whiskeytown start month: %s; Start index: %d"%(whiskeytown_start_month, whiskeytown_start_index)
	for line in ops_data["Whiskeytown"][1:]:
		# print "Passing line to CVP.make_ops_tsc: %s"%(line)
		data_month = whiskeytown_start_month
		data_year = target_year
		try:
			early_val = float(line.split(',')[whiskeytown_start_index - 1].strip())
			data_month = CVP.month_TLA[CVP.last_month(CVP.month_index(whiskeytown_start_month))]
			if data_month == "DEC":
				data_year -= 1
		except:
			pass
		# print "Start_index = %d\nData_Month = %s"%(start_index, data_month)
		# print "Passing line to CVP.make_ops_tsc: %s"%(line)
		whiskeytown_ts_list.append(CVP.make_ops_tsc("Whiskeytown", data_year, data_month, line, ops_label=ops_import_F_part))
		# ts_count += 1

	pattern_DSS_file_name = ""
	pattern_path = ""

	with open(flow_pattern_config_file) as infile:
		pattern_config_str = infile.read()
	pattern_config_str = re.sub(r"<!--.*?-->", '', pattern_config_str).strip()
	pattern_config_str = re.sub("\n\n", "\n", pattern_config_str)
	print pattern_config_str

	for line in pattern_config_str.split('\n'):
		token = line.strip().split(',')
		if len(token) != 3:
			print "File %s line \n\t \"%s\"\nis not a valid ID for a position analysis DSS record."%(flow_pattern_config_file,line)
			continue
		if line.split(',')[0].strip().upper() == "SHASTA":
			pattern_DSS_file_name = line.split(',')[1].strip().strip('\\')
			pattern_path = line.split(',')[2].strip()
	if len(pattern_DSS_file_name) == 0 or len(pattern_path) == 0:
		print "Error reading flow pattern configuration file\n\t%s"%(flow_pattern_config_file)
		print "Shasta pattern DSS file or path not found."
		return 0
	if not os.path.isabs(pattern_DSS_file_name):
		pattern_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), pattern_DSS_file_name)
		# print "Flow pattern for Shasta in \n\t%s"%(pattern_DSS_file_name)
		# print "\t" + pattern_path

	met_DSS_file_name = ""
	airtemp_path = ""
	with open(DSS_map_file) as infile:
		for line in infile:
			if (line.split(',')[0].strip() == "ReddingAirport-RED" and
				line.split(',')[1].strip() == "Air Temperature"):
				met_DSS_file_name = line.split(',')[2].strip().strip('\\')
				airtemp_path = line.split(',')[3].strip()
	if len(met_DSS_file_name) == 0 or len(airtemp_path) == 0:
		print "Error reading Shasta air temperature data configuration from file\n\t%s"%(DSS_map_file)
		print "Air temperature DSS file or path not found."
		return 0
	if not os.path.isabs(met_DSS_file_name):
		met_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), met_DSS_file_name)


	outDSS = hec.heclib.dss.HecDss.open(BC_output_DSS_filename)
	patternDSS = hec.heclib.dss.HecDss.open(pattern_DSS_file_name)
	temperatureDSS = hec.heclib.dss.HecDss.open(met_DSS_file_name)

	tsm_list = []
	tsmath_acc_dep = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(start_time.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_acc_dep.setUnits("CFS")
	tsmath_acc_dep.setType("PER-AVER")
	tsmath_acc_dep.setTimeInterval("1DAY")
	tsmath_acc_dep.setLocation("SHASTA")
	tsmath_acc_dep.setParameterPart("FLOW-ACC-DEP")
	tsmath_acc_dep.setVersion(BC_F_part)
	for ts in shasta_ts_list:
		print "TS Parameter = %s"%(ts.parameter.upper())
		if ts.parameter.upper() == "INFLOW":
			tsmath_flow_monthly = tsmath(ts)
			tsm_list.append(tsmath_flow_monthly)
			print "reading pattern from file: " + pattern_DSS_file_name
			print "\t" + pattern_path
			tsmath_pattern = patternDSS.read(pattern_path)
			tsmath_weighted = CVP.weight_transform_monthly_to_daily(tsmath(ts), tsmath_pattern)
			tsmath_weighted.setPathname(ts.fullName)
			tsmath_weighted.setTimeInterval("1DAY")
			tsmath_weighted.setParameterPart("FLOW-IN")
			tsmath_weighted.setVersion(BC_F_part)
			tsm_list.append(tsmath_weighted)
		elif ts.parameter.upper() == "EST. EVAP.":
			tsmath_evap_monthly = tsmath(ts)
			tsm_list.append(tsmath_evap_monthly)
			tsmath_acc_dep = tsmath_acc_dep.subtract(
				CVP.uniform_transform_monthly_to_daily(tsmath(ts)))
		elif ts.parameter.upper() == "TOTAL SHASTA RELEASE":
			tsmath_release_monthly = tsmath(ts)
			tsm_list.append(tsmath_release_monthly)
			tsmath_release = CVP.uniform_transform_monthly_to_hourly(tsmath(ts))
			tsmath_release.setPathname(ts.fullName)
			tsmath_release.setTimeInterval("1HOUR")
			tsmath_release.setParameterPart("FLOW-RELEASE")
			tsmath_release.setVersion(BC_F_part)
			tsm_list.append(tsmath_release)
		else:
			tsm_list.append(tsmath(ts))
	tsm_list.append(tsmath_acc_dep)
	for ts in whiskeytown_ts_list:
		print "TS Parameter = %s"%(ts.parameter.upper())
		if ts.parameter.upper() == "SPRING CR.":
			tsmath_sp_cr_monthly = tsmath(ts)
			tsm_list.append(tsmath_sp_cr_monthly)
			tsmath_sp_cr = CVP.uniform_transform_monthly_to_hourly(tsmath(ts))
			tsmath_sp_cr.setPathname(ts.fullName)
			tsmath_sp_cr.setTimeInterval("1HOUR")
			tsmath_sp_cr.setParameterPart("FLOW-PP")
			tsmath_sp_cr.setVersion(BC_F_part)
			tsm_list.append(tsmath_sp_cr)

	tsmath_zero_flow_day = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(start_time.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_zero_flow_day.setUnits("CFS")
	tsmath_zero_flow_day.setType("PER-AVER")
	tsmath_zero_flow_day.setTimeInterval("1DAY")
	tsmath_zero_flow_day.setLocation("ZERO-BY-DAY")
	tsmath_zero_flow_day.setParameterPart("FLOW-ZERO")
	tsmath_zero_flow_day.setVersion(BC_F_part)
	tsm_list.append(tsmath_zero_flow_day)

	tsmath_zero_flow_hour = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(start_time.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1HOUR", "0M", 0.0)
	tsmath_zero_flow_hour.setUnits("CFS")
	tsmath_zero_flow_hour.setType("PER-AVER")
	tsmath_zero_flow_hour.setTimeInterval("1Hour")
	tsmath_zero_flow_hour.setLocation("ZERO-BY-HOUR")
	tsmath_zero_flow_hour.setParameterPart("FLOW-ZERO")
	tsmath_zero_flow_hour.setVersion(BC_F_part)
	tsm_list.append(tsmath_zero_flow_hour)


	# Table of tributary weights by month
	'''tributary_weights = {
		"Shasta-Sac-in":(0.176691, 0.197511, 0.20499, 0.219252, 0.208572, 0.150637, 0.094333, 0.076416, 0.068915, 0.081511, 0.10148, 0.161356),
		"Shasta-McCloud-in":(0.121142, 0.130348, 0.126563, 0.114731, 0.102656, 0.101719, 0.102699, 0.099085, 0.10394, 0.115341, 0.109092, 0.131823),
		"Shasta-Sulanharas-in":(0.026955, 0.032142, 0.034431, 0.035416, 0.031629, 0.019334, 0.009611, 0.007116, 0.006271, 0.008286, 0.011383, 0.02422),
		"Shasta-Pit-in":(0.675211, 0.639999, 0.634016, 0.630601, 0.657144, 0.72831, 0.793356, 0.817382, 0.820874, 0.794863, 0.778044, 0.682601)}
	'''
	tributary_weights = {
		"Shasta-Sac-in":(0.212770745, 0.224327192, 0.221179858, 0.231031865, 0.22998096, 0.174508497, 0.096498474, 0.074162081, 0.066134982, 0.085930713, 0.110001981, 0.208573952),
		"Shasta-McCloud-in":(0.138567582, 0.157547951, 0.139190927, 0.129798785, 0.107066929, 0.097430013, 0.099133208, 0.094616182, 0.097972639, 0.111942455, 0.109353393, 0.151801944),
		"Shasta-Sulanharas-in":(0.037029058, 0.042679679, 0.040961806, 0.039603204, 0.037053932, 0.024035906, 0.01008518, 0.006934946, 0.006026154, 0.009676144, 0.013545787, 0.038994156),
		"Shasta-Pit-in":(0.611632586, 0.575445235, 0.598667383, 0.599566102, 0.625898182, 0.704025567, 0.794283211, 0.824286819, 0.82986623, 0.792450666, 0.767098904, 0.600629926)}
	names_flows = {}
	for tsm in CVP.split_time_series_monthly(tsmath_weighted, tributary_weights, "FLOW-IN"):
		tsm.setVersion(BC_F_part)
		tsm_list.append(tsm)
		names_flows[tsm.getContainer().location] = tsm



	#River, Intercept (deg C), Flow Coef (cfs), Air Temp Coef (deg C), RMS Error (deg C)
	tributary_temp_regression_coefficients = {
		"Shasta-Sac-in": (1.1597557, -2.5038779e-04, 0.62590134, 1.6474143),
		"Shasta-Pit-in": (3.2822256, -1.541817e-04, 0.55336446, 1.4528962),
		"Shasta-McCloud-in": (1.735364, 2.1436048e-04, 0.48995328, 1.1855532)}
	tsmath_airtemp = temperatureDSS.read(airtemp_path)
	for key in tributary_temp_regression_coefficients.keys():
		tsm = CVP.evaluate_temp_regression(names_flows[key], tsmath_airtemp, tributary_temp_regression_coefficients[key])
		tsm.setVersion(BC_F_part)
		tsm_list.append(tsm)

	for tsmath_item in tsm_list:
		rv_lines.append("%s,%s,%s,%s"%(
			tsmath_item.getContainer().location, tsmath_item.getContainer().parameter,
			Project.getCurrentProject().getRelativePath(BC_output_DSS_filename),
			tsmath_item.getContainer().fullName))
		outDSS.write(tsmath_item)

	outDSS.done()
	patternDSS.done()
	temperatureDSS.done()

	return rv_lines


'''
Imports a CVP ops spreadsheet saved as comma-separated values
Returns a dictionary with keys that match the list of forecast locations in the second argrument
Dictionary values are lists of CSV lines that "belong" to the location named in the key
'''
def import_CVP_Ops_csv(ops_fname, forecast_locations):
	current_location = None
	start_month = None
	first_date_index = -1
	location_count = 0
	ts_count = 0
	data_lines = []
	rv_dictionary = {}
	calendar = ""

	with open(ops_fname) as infile:
		num_lines = 0; num_data_lines = 0
		for line in infile:
			num_lines += 1
			line_contains_months = False
			token = line.split(',')
			# figure out what columns our data start in, what month we're looking at, and ignore blank lines
			# the sample spreadsheet had an unused summary block starting in column AA, which I'm ignoring
			num_t = 0; num_val = 0
			for t in token[:26]:
				if len(t.strip()) > 0:
					num_val += 1
					if not line_contains_months and t.strip().upper() in CVP.month_TLA:
						line_contains_months = True
						first_date_index = num_t
						start_month = t.strip().upper()
						# print "Calendar line %s: "%(line)
						# print "Found \"%s\" in column %d"%(t.strip(), num_t + 1)
						calendar = line
				num_t += 1
			if num_val == 0:
				continue # don't include this line in the result

			if token[0].strip() in forecast_locations and len(calendar) > 0:
				if location_count > 0:
					rv_dictionary[current_location] = data_lines
					data_lines = []
				current_location = token[0].strip()
				print "setting current location to %s"%(current_location)
				data_lines.append("%d,%s"%(first_date_index, calendar))
				location_count += 1
				calendar = ""
				continue

			if not line_contains_months:
				data_lines.append(line)
				ts_count += 1

	rv_dictionary[current_location] = data_lines #
	print "Found %d forecast locations and %d time series in ops file \n\t%s."%(
		location_count, ts_count, ops_fname)
	return rv_dictionary


def monthFromDateStr(str):
	month_TLA = ["NM", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
	for token in str.split():
		if token.strip().upper() in month_TLA:
			return token.strip().upper()
	return None

'''
Imports a CVP ops spreadsheet saved as XLS or XLSX format
Returns a dictionary with keys that match the list of forecast locations in the second argrument
Dictionary values are lists of CSV lines that "belong" to the location named in the key

Excel formats are decoded by the Apache POI library. See import block at the top of the
file. The instructional web sites below helped with interpreting values from formula cells
https://www.baeldung.com/java-apache-poi-cell-string-value
https://www.baeldung.com/java-read-dates-excel
'''
def import_CVP_Ops_xls(ops_fname, forecast_locations, sheet_number=0):
	current_location = None
	start_month = None
	first_date_index = -1
	location_count = 0
	ts_count = 0
	data_lines = []
	rv_dictionary = {}
	calendar = ""

	try:
		if ops_fname.endswith(".xlsx"):
			workbook = XSSFWorkbook(
				java.io.FileInputStream(java.io.File(ops_fname)))
		if ops_fname.endswith(".xls"):
			workbook = HSSFWorkbook(
				java.io.FileInputStream(java.io.File(ops_fname)))
	except Exception as e:
		raise e

	sheet = workbook.getSheetAt(sheet_number)
	formatter = SSUsermodel.DataFormatter(True)
	num_lines = 0; num_data_lines = 0
	for row in sheet.iterator():
		num_lines += 1
		line_contains_months = False
		token = []
		for cell in row.cellIterator():
			# This business -- Cell.CELL_TYPE_XXX -- has been revised a couple of times
			# between POI version 3.8 and 4.x. Watch out it doesn't bite us
			if cell.getCellType() == SSUsermodel.Cell.CELL_TYPE_FORMULA:
				cachedType = cell.getCachedFormulaResultType()
				# print str(cachedType) + " : " + formatter.formatCellValue(cell)
				if cachedType == SSUsermodel.Cell.CELL_TYPE_NUMERIC:
					if SSUsermodel.DateUtil.isCellDateFormatted(cell):
						token.append(monthFromDateStr(str(cell.getDateCellValue())))
					else:
						token.append(str(cell.getNumericCellValue()))
				if cachedType == SSUsermodel.Cell.CELL_TYPE_STRING:
					token.append(str(cell.getStringCellValue()))
			else:
				token.append(formatter.formatCellValue(cell))
		# figure out what columns our data start in, what month we're looking at, and ignore blank lines
		num_t = 0; num_val = 0
		for t in token:
			if len(t.strip()) > 0:
				num_val += 1
				# if there's a month label in the first 6 cells of the row,
				# it's a calendar line
				if ((not line_contains_months) and
					num_t < 6 and
					t.strip().upper() in CVP.month_TLA):
					line_contains_months = True
					first_date_index = num_t
					start_month = t.strip().upper()
					print "Calendar line %d: "%(num_lines)
					print "Found \"%s\" in column %d"%(t.strip(), num_t + 1)
					calendar = ','.join(token)
			num_t += 1
		if num_val == 0:
			continue # don't include this row in the result

		if token[0].strip() in forecast_locations and len(calendar) > 0:
			if location_count > 0:
				rv_dictionary[current_location] = data_lines
				data_lines = []
			current_location = token[0].strip()
			print "setting current location to %s"%(current_location)
			data_lines.append("%d,%s"%(first_date_index, calendar))
			location_count += 1
			calendar = ""
			continue

		if not line_contains_months:
			data_lines.append(','.join(token))
			ts_count += 1

	rv_dictionary[current_location] = data_lines #
	print "Found %d forecast locations and %d time series in ops file \n\t%s."%(
		location_count, ts_count, ops_fname)
	return rv_dictionary
