# Standard-library imports and HEC/ResSim project interfaces used by the script.
import os, sys
import re
from com.rma.io import DssFileManagerImpl
from com.rma.model import Project

# HEC-DSS and HEC time-series classes provide storage and time-series operations.
import hec.heclib.dss
import hec.heclib.util.HecTime as HecTime
import hec.io.TimeSeriesContainer as tscont
import hec.hecmath.TimeSeriesMath as tsmath
from hec.script import MessageBox

# Forecast configuration utilities provide project-specific configuration paths.
import usbr.wat.plugins.actionpanel.model.forecast as fc

# Make project forecast scripts importable, then load the CVP operations utilities.
sys.path.append(os.path.join(Project.getCurrentProject().getWorkspacePath(), "forecast", "scripts"))

import CVP_ops_tools as CVP
reload(CVP)

DEBUG = True

def build_BC_data_sets(AP_start_time, AP_end_time, BC_F_part, BC_output_DSS_filename, ops_file_name, DSS_map_filename,
		position_analysis_year=None,
		position_analysis_config_filename=None,
		met_F_part=None,
		met_output_DSS_filename=None,
		flow_pattern_config_filename=None,
		ops_import_F_part=None):
	"""Build meteorological, hydrologic, and water-temperature boundary data sets.

	Parameters
	----------
	AP_start_time : HecTime
		Start of the simulation-group run period.
	AP_end_time : HecTime
		End of the simulation-group run period.
	BC_F_part : str
		DSS F-part assigned to generated boundary-condition records.
	BC_output_DSS_filename : str
		DSS file receiving generated boundary-condition time series.
	ops_file_name : str
		CVP operations spreadsheet or CSV file used as operational input.
	DSS_map_filename : str
		File receiving the location, parameter, DSS file, and pathname map.
	position_analysis_year : int, optional
		Historical source year used for meteorological positional analysis.
	position_analysis_config_filename : str, optional
		Configuration describing source meteorological DSS records.
	met_F_part : str, optional
		DSS F-part for meteorological records; defaults to ``BC_F_part``.
	met_output_DSS_filename : str, optional
		Separate meteorological DSS output file, if required.
	flow_pattern_config_filename : str, optional
		Configuration describing flow-pattern records used for disaggregation.
	ops_import_F_part : str, optional
		Label applied to records imported from the operations input file.

	Returns
	-------
	int
		Number of meteorological and operational map records generated, or 0
		when operational boundary-condition generation fails.

	Notes
	-----
	Relative paths are resolved against the current project workspace. The
	routine writes DSS records and a location/path mapping file as side effects.
	"""

	# Resolve project-relative files and establish defaults for optional DSS
	# metadata and configuration inputs.
	if not os.path.isabs(BC_output_DSS_filename):
		BC_output_DSS_filename = os.path.join(Project.getCurrentProject().getWorkspacePath(), BC_output_DSS_filename)
	if not os.path.isabs(ops_file_name):
		ops_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), ops_file_name)
	if not met_F_part:
		met_F_part = BC_F_part
	if not ops_import_F_part:
		ops_import_F_part = os.path.basename(ops_file_name)

	# Meteorological output defaults to the boundary-condition DSS file unless
	# the caller supplies a separate destination.
	if not met_output_DSS_filename:
		met_output_DSS_filename=BC_output_DSS_filename
	elif not os.path.isabs(met_output_DSS_filename):
		met_output_DSS_filename = os.path.join(Project.getCurrentProject().getWorkspacePath(), met_output_DSS_filename)

	# Use project forecast configuration defaults when explicit configuration
	# filenames were not supplied.
	if not position_analysis_config_filename:
		position_analysis_config_filename = fc.ForecastConfigFiles.getHistoricalMetFile()
	elif not os.path.isabs(position_analysis_config_filename):
		position_analysis_config_filename = os.path.join(Project.getCurrentProject().getWorkspacePath(), position_analysis_config_filename)
	if not flow_pattern_config_filename:
		flow_pattern_config_filename = fc.ForecastConfigFiles.getFlowPatternFile()
	elif not os.path.isabs(flow_pattern_config_filename):
		flow_pattern_config_filename = os.path.join(Project.getCurrentProject().getWorkspacePath(), flow_pattern_config_filename)
	if not os.path.isabs(DSS_map_filename):
		DSS_map_filename = os.path.join(Project.getCurrentProject().getWorkspacePath(), DSS_map_filename)

	# Report the principal processing inputs and output destinations.
	print "\n########"
	print "\tGenerating Boundary Conditions for Shasta/Trinity models"
	print "########\n"

	print "CVP Ops Data file: %s"%ops_file_name
	print "Met data config file: %s"%position_analysis_config_filename
	print "Flow pattern config file: %s"%flow_pattern_config_filename
	print "Boundary Condition output DSS file: %s"%BC_output_DSS_filename
	print "Met data output DSS file: %s"%met_output_DSS_filename
	print "Location/Path map file: %s"%DSS_map_filename

	print "\nPreparing Meteorological Data..."

	# Generate positional-analysis meteorological records over the requested
	# analysis period.
	met_lines = create_positional_analysis_met_data(AP_start_time.year(), position_analysis_year, AP_start_time, AP_end_time,
		position_analysis_config_filename, met_output_DSS_filename, met_F_part)

	# Initialize the DSS map with the meteorological records generated above.
	with open(os.path.join(Project.getCurrentProject().getWorkspacePath(), DSS_map_filename), "w") as mapfile:
		mapfile.write("location,parameter,dss file,dss path\n")
		for line in met_lines:
			mapfile.write(line + '\n')
			if DEBUG: print(line)

	print("Met process complete.\n\nPreparing hydro and WC boundary conditions...")

	# Generate hydrologic and water-temperature boundary conditions from the
	# operations data and append their map entries if processing succeeds.
	ops_lines = create_ops_BC_data(ops_file_name, AP_start_time, AP_end_time,
		BC_output_DSS_filename, BC_F_part, ops_import_F_part, flow_pattern_config_filename, DSS_map_filename)
	if not ops_lines:
		return 0

	with open(os.path.join(Project.getCurrentProject().getWorkspacePath(), DSS_map_filename), "a") as mapfile:
		for line in ops_lines:
			mapfile.write(line)
			mapfile.write('\n')
	print "\nBoundary condition report written to: %s\n"%(DSS_map_filename)

	return len(met_lines) + len(ops_lines)


def create_positional_analysis_met_data(target_year, source_year, start_time, end_time,
position_analysis_config_filename, met_output_DSS_filename, met_F_part):
	"""Shift historical meteorological data into a target analysis year.

	This is a simple time-shifter for meteorological positional-analysis data.
	Location-specific information and DSS file/path combinations are supplied
	through the positional-analysis configuration file rather than being
	hard-coded in this function.

	Parameters
	----------
	target_year : int
		Year to which the meteorological sequence is shifted.
	source_year : int
		Historical year supplying the meteorological sequence.
	start_time : HecTime
		Start of the target analysis period.
	end_time : HecTime
		End of the target analysis period.
	position_analysis_config_filename : str
		Configuration describing source and destination DSS records.
	met_output_DSS_filename : str
		DSS file receiving the shifted meteorological records.
	met_F_part : str
		DSS F-part assigned to shifted records.

	Returns
	-------
	list of str
		CSV-formatted mapping records for the generated meteorological series.

	Notes
	-----
	The source and generated series are assumed to have compatible time steps.
	"""
	print "Calculating positional met data..."
	print "Historical Met File: %s"%(fc.ForecastConfigFiles.getHistoricalMetFile())
	print "Position Analysis Met File: %s"%position_analysis_config_filename
	diff_years = target_year - source_year
	print "Shifting met data from %d to %d (%d years)."%(source_year, target_year, diff_years)

	rv_lines = []
	met_config_str = ""
	print "Met output DSS file: %s"%(met_output_DSS_filename)
	met_config_lines = getConfigLines(position_analysis_config_filename)

	# Process each configured meteorological source/destination mapping after
	# the configuration header.
	for line in met_config_lines[1:]:
		token = line.strip().split(',')
		dest_count = 0
		try:
			dest_count = int(token[4])
		except:
			print "File %s line \n\t \"%s\"\nis not a valid ID for a position analysis DSS record."%(position_analysis_config_filename,line)
			print "Can't read an integer value from \"%s\"."%(token[4])
			continue

		# Validate the expected number of fields from the configured destination
		# count before accessing DSS pathname elements.
		target_line_length = 5 + 2*dest_count
		if len(token) != target_line_length:
			print "File %s line \n\t \"%s\"\nis not a valid ID for a position analysis DSS record."%(position_analysis_config_filename,line)
			continue

		# Resolve and open the configured historical meteorological DSS source.
		#source_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), token[0].strip('\\'))
		source_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), token[2].strip().strip('\\'))
		ts_read = hec.heclib.dss.HecTimeSeries()
		ts_read.setDSSFileName(source_DSS_file_name)
		if DEBUG: print "Reading %s from DSS file %s."%(token[3].strip(), source_DSS_file_name)

		# Reconstruct source and destination DSS paths while intentionally
		# leaving the D-part blank.
		source_path_parts = token[3].strip().strip('/').split('/', 5)
		dest_path_parts = token[6].strip().strip('/').split('/', 5)
		source_path = dest_path = '/'
		for index in (0,1,2,4,5):
			source_path += (source_path_parts[index] + '/')
			dest_path += (dest_path_parts[index] + '/')
			if index == 2:
				source_path += '/'
				dest_path += '/'

		# Read the historical source series identified by the reconstructed path.
		tsc_source = tscont()
		tsc_source.fullName = source_path
		status = ts_read.read(tsc_source, False)
		if status < 0:
			print "Failed to read meteorologic time series %s \n\tfrom DSS file %s"%(source_path, source_DSS_file_name)
			ts_read.done()
			continue

		tsmath_source = tsmath(tsc_source)
		time_step_label = token[3].strip().split('/')[5]
		if DEBUG:  print "\tTime series contains %d values."%(tsmath_source.getContainer().numberValues)
		if DEBUG:  print "\tShifting time series with shiftInTime(%s)."%("%dMo"%(diff_years*12))
		# tsmath_shift = tsmath_source.shiftInTime("%dYrar"%(diff_years))

		# Build the target time grid with one day of end padding, then locate
		# the corresponding start position in the historical source year.
		padded_end_time = HecTime()
		padded_end_time.set(end_time.value() + 1440)
		tsmath_shift = tsmath.generateRegularIntervalTimeSeries(
			"%s 0000"%(start_time.date(4)),
			"%s 2400"%(padded_end_time.date(4)),
			time_step_label, "0M", 1.0)
		time_seek = HecTime(tsmath_shift.firstValidDate(), HecTime.MINUTE_INCREMENT)
		time_seek.setYearMonthDay(time_seek.year() - diff_years, time_seek.month(), time_seek.day(), time_seek.minutesSinceMidnight())

		# Abort this positional-analysis operation if the requested shifted
		# period begins before the available historical record.
		if time_seek.getMinutes() < tsmath_source.firstValidDate():
			print "Met position time shift out of range at source start..."
			return ['']

		source_container = tsmath_source.getContainer()
		shift_container = tsmath_shift.getContainer()
		start_index = 0

		# Find the first historical value corresponding to the shifted target
		# start time.
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

		# Verify that the generated container retains the expected number of
		# values after copying the historical sequence.
		if len(shift_container.values) != shift_container.numberValues:
			print "You doofus!\nlen(values)=%d\nnumberValues=%d\n"%(len(shift_container.values), shift_container.numberValues)
			return ['']

		# Preserve source units/type and assign the destination DSS pathname
		# metadata before writing the shifted record.
		tsmath_shift.setType(tsmath_source.getType())
		tsmath_shift.setUnits(tsmath_source.getUnits())
		tsmath_shift.setPathname(dest_path)
		tsmath_shift.setVersion(met_F_part)
		ts_write = hec.heclib.dss.HecTimeSeries()
		ts_write.setDSSFileName(met_output_DSS_filename)
		if DEBUG: print "Writing %s to DSS file %s."%(shift_container.fullName, met_output_DSS_filename)
		ts_write.write(tsmath_shift.getData())
		ts_write.done()
		ts_read.done()

		# Record the generated DSS location/path mapping for downstream use.
		#met_loc, met_param = token[1].strip().split('<', 1)
		met_loc = token[0]
		met_param = token[1]
		rv_lines.append("%s,%s,%s,%s"%(met_loc.strip(), met_param.strip().strip('>'),
		Project.getCurrentProject().getRelativePath(met_output_DSS_filename),
		tsmath_shift.getContainer().fullName))

	return rv_lines

def shift_monthly_averages(source_tsm, AP_start_time, AP_end_time):
	"""Repeat monthly average values over the requested analysis period.

	Parameters
	----------
	source_tsm : TimeSeriesMath
		Source monthly-average time series used as the repeating pattern.
	AP_start_time : HecTime
		Start of the requested analysis period.
	AP_end_time : HecTime
		End of the requested analysis period.

	Returns
	-------
	TimeSeriesMath
		Monthly time series spanning the requested period.

	Notes
	-----
	The source data are expected to span complete years so indexing can wrap
	from the end of the source sequence back to its beginning.
	"""

	# Copy the requested bounds so the end-of-month adjustments below do not
	# modify the HecTime objects supplied by the caller.
	shifted_start_time = HecTime()
	shifted_start_time.set(AP_start_time)
	shifted_end_time = HecTime()
	shifted_end_time.set(AP_end_time)

	# Move both analysis bounds to 2400 at the end of their respective months.
	for hec_time in (shifted_start_time, shifted_end_time):
		hec_time.setTime("2400")
		hec_time.addDays(CVP.get_days_in_month(hec_time.month(),hec_time.year()) - hec_time.day())

	# Generate the target monthly series and carry forward the source units,
	# data type, location, and parameter metadata.
	rv_tsmath = tsmath.generateRegularIntervalTimeSeries(shifted_start_time.date(8), shifted_end_time.date(8), "1MON", 1.0)
	rv_tsmath.setUnits(source_tsm.getUnits())
	rv_tsmath.setType(source_tsm.getType())
	rv_tsmath.setLocation(source_tsm.getContainer().location)
	rv_tsmath.setParameterPart(source_tsm.getContainer().parameter)

	# Find the source-series position corresponding to the target starting month.
	seek_index = 0
	seek_time = HecTime()
	seek_time.set(source_tsm.getContainer().times[seek_index])
	while seek_time.month() < shifted_start_time.month():
		seek_index += 1
		seek_time.set(source_tsm.getContainer().times[seek_index])

	# Copy monthly values into the target period, wrapping to the start of the
	# source sequence after its final value.
	dest_index = 0
	while dest_index < rv_tsmath.getContainer().numberValues:
		rv_tsmath.getContainer().values[dest_index] = source_tsm.getContainer().values[seek_index]
		dest_index += 1
		seek_index += 1
		# wrap around to the beginning of the source when you hit the end
		# not that this presumems that the source data set spans whole years
		if seek_index >= source_tsm.getContainer().numberValues: seek_index = 0

	#return the time-series math object
	return rv_tsmath


def getConfigLines(fileName):
	"""Read a configuration file and remove supported comment syntax.

	Parameters
	----------
	fileName : str
		Path to the configuration file.

	Returns
	-------
	list of str
		Non-comment configuration content split into individual lines.
	"""
	# Remove both XML-style comments and hash comments before collapsing
	# repeated blank lines in the configuration text.
	commentRE = re.compile(r"<!--.*?-->", re.S)
	hashCommentRE = re.compile(r"#.*")
	with open(fileName) as infile:
		config_str = infile.read()
	config_str = commentRE.sub("", config_str)
	config_str = hashCommentRE.sub("", config_str).strip()
	config_str = re.sub(r"\n+", "\n", config_str)
	return  config_str.split('\n')


def create_ops_BC_data(ops_file_name, start_time, end_time, BC_output_DSS_filename,
	BC_F_part, ops_import_F_part, flow_pattern_config_filename, DSS_map_filename):
	"""Create hydrologic and water-temperature boundary-condition records.

	Processes the contents of the CVP operations spreadsheet into flow and
	water-temperature boundary conditions.

	Parameters
	----------
	ops_file_name : str
		CVP operations spreadsheet or CSV file.
	start_time : HecTime
		Start of the forecast time window.
	end_time : HecTime
		End of the forecast time window.
	BC_output_DSS_filename : str
		DSS file receiving generated boundary-condition records.
	BC_F_part : str
		DSS F-part assigned to generated records.
	ops_import_F_part : str
		Version label associated with imported operations data.
	flow_pattern_config_filename : str
		Configuration identifying flow-pattern DSS records.
	DSS_map_filename : str
		Location/path map containing meteorological DSS references.

	Returns
	-------
	list of str or None
		CSV-formatted map records for generated DSS time series, or ``None``
		when required input data or configuration cannot be processed.

	Notes
	-----
	This routine imports CVP operations data, disaggregates monthly quantities,
	constructs reservoir water balances, estimates tributary temperatures, and
	writes the resulting time series to DSS.
	"""
	print "Processing boundary conditions for upper Sacramento River from ops file:\n\t%s"%(ops_file_name)
	print "  Forecast time window start: %s"%(start_time.dateAndTime(4))
	print "  Forecast time window end: %s"%(end_time.dateAndTime(4))

	# Define the spreadsheet locations recognized by the parser and the subset
	# that supplies profile-date information for active boundary processing.
	forecast_locations = ["Trinity/Clair Engle", "Whiskeytown", "Shasta", "Oroville", "Folsom", "New Melones", " SAN LUIS/O'NEILL", "DELTA"]
	active_locations = ["Trinity/Clair Engle", "Whiskeytown", "Shasta"]

	rv_lines = []

	# Select the appropriate operations-data importer from the input extension.
	if ops_file_name.endswith(".xls") or ops_file_name.endswith(".xlsx"):
		try:
			ops_data = CVP.import_CVP_Ops_xls(ops_file_name, forecast_locations, active_locations)
		except Exception as e:
			print "Failed to read operations file:%s"%ops_file_name
			print "\t%s"%str(e)
			return None
	else:
		ops_data = CVP.import_CVP_Ops_csv(ops_file_name, forecast_locations, active_locations)

	# Extract the profile date from the active-location records and remove that
	# metadata row before interpreting the remaining rows as time series.
	profile_date = None

	for key in ops_data.keys():
		if DEBUG:
			print "ops_data key: %s"%(key)
		if ops_data[key][1].strip().upper().startswith("PROFILEDATE"):
			profile_date = ops_data[key][1].split(':')[1].strip()
			del ops_data[key][1]

	# Normalize the profile date to the compact form expected by HecTime.
	if profile_date:
		try:
			date_parts = profile_date.split('-', 2)
			if len(date_parts[2]) < 4: date_parts[2] = "20" + date_parts[2]
			profile_date = "%s%s%s"%(date_parts[0],date_parts[1],date_parts[2])
		except Exception as e:
			print "Failed to read profile date from string:%s"%profile_date
			print "\t%s"%str(e)
			return None
		print "Profile date: %s"%profile_date

	# Read the Shasta calendar metadata that establishes the starting column and
	# month for monthly values in the operations spreadsheet.
	shasta_tsc_list = []
	shasta_calendar = ops_data["Shasta"][0].split(',')
	shasta_start_index = int(shasta_calendar[0])
	shasta_start_month = shasta_calendar[shasta_start_index + 1].strip().upper()
	if DEBUG: print "\n Shasta start month: %s; Start index: %d"%(shasta_start_month, shasta_start_index)

	# Establish the operations start date and, when a profile date is present,
	# the number of days represented by the partial first month.
	ops_start_date = HecTime()
	days_in_first_month = None
	if profile_date:
		ops_start_date.set(profile_date, "2400")
		days_in_first_month = 1 + CVP.get_days_in_month(CVP.month_index(shasta_start_month), ops_start_date.year()) - ops_start_date.day()
	else:
		ops_start_date.set("01%s%d"%(shasta_start_month, start_time.year()), "0001")
		if ops_start_date > start_time:
			ops_start_date.set("01%s%d"%(shasta_start_month, start_time.year()-1), "0001")

	# Convert each Shasta spreadsheet row into a monthly TimeSeriesContainer,
	# adjusting the starting month when a numeric value precedes the nominal
	# calendar start column.
	for line in ops_data["Shasta"][1:]:
		data_month = shasta_start_month
		data_year = ops_start_date.year()
		if len(line.split(',')[0]) ==0:
			continue
		if CVP.is_convertable_to_float(line.split(',')[shasta_start_index - 1].strip()):
			data_month = CVP.month_TLA[CVP.previous_month(CVP.month_index(shasta_start_month))]
			if data_month == "DEC":
				data_year -= 1
		if DEBUG: print "Start_index = %d\nData_Month = %s"%(shasta_start_index, data_month)
		if DEBUG: print "Passing line to CVP.make_ops_tsc: %s"%(line)
		shasta_tsc_list.append(CVP.make_ops_tsc("SHASTA", data_year, data_month, line, ops_label=ops_import_F_part))

	# Apply the same calendar interpretation to the Whiskeytown operations
	# records.
	whiskeytown_tsc_list = []
	whiskeytown_calendar = ops_data["Whiskeytown"][0].split(',')
	whiskeytown_start_index = int(whiskeytown_calendar[0])
	whiskeytown_start_month = whiskeytown_calendar[whiskeytown_start_index + 1].strip().upper()
	if DEBUG: print "\n Whiskeytown start month: %s; Start index: %d"%(whiskeytown_start_month, whiskeytown_start_index)
	for line in ops_data["Whiskeytown"][1:]:
		data_month = whiskeytown_start_month
		data_year = ops_start_date.year()
		if len(line.split(',')[0]) ==0:
			continue
		if CVP.is_convertable_to_float(line.split(',')[whiskeytown_start_index - 1].strip()):
			data_month = CVP.month_TLA[CVP.previous_month(CVP.month_index(whiskeytown_start_month))]
			if data_month == "DEC":
				data_year -= 1
		if DEBUG: print "Start_index = %d\nData_Month = %s"%(whiskeytown_start_index, data_month)
		if DEBUG: print "Passing line to CVP.make_ops_tsc: %s"%(line)
		whiskeytown_tsc_list.append(CVP.make_ops_tsc("Whiskeytown", data_year, data_month, line, ops_label=ops_import_F_part))

	# Build the corresponding monthly containers for Trinity/Clair Engle.
	trinity_tsc_list = []
	trinity_calendar = ops_data["Trinity/Clair Engle"][0].split(',')
	trinity_start_index = int(trinity_calendar[0])
	trinity_start_month = trinity_calendar[trinity_start_index + 1].strip().upper()
	if DEBUG: print "\n Trinity start month: %s; Start index: %d"%(trinity_start_month, trinity_start_index)
	for line in ops_data["Trinity/Clair Engle"][1:]:
		data_month = trinity_start_month
		data_year = ops_start_date.year()
		if len(line.split(',')[0]) ==0:
			continue
		if CVP.is_convertable_to_float(line.split(',')[trinity_start_index - 1].strip()):
			data_month = CVP.month_TLA[CVP.previous_month(CVP.month_index(trinity_start_month))]
			if data_month == "DEC":
				data_year -= 1
		if DEBUG: print "Start_index = %d\nData_Month = %s"%(trinity_start_index, data_month)
		if DEBUG: print "Passing line to CVP.make_ops_tsc: %s"%(line)
		trinity_tsc_list.append(CVP.make_ops_tsc("Trinity/Clair Engle", data_year, data_month, line, ops_label=ops_import_F_part))

	# Load flow-pattern configuration and the existing DSS map so subsequent
	# reservoir processing can locate pattern and meteorological records.
	shasta_pattern_DSS_file_name = whiskeytown_pattern_DSS_file_name = trinity_pattern_DSS_file_name =""
	shasta_pattern_path = whiskeytown_pattern_path = trinity_pattern_path = None
	flow_pattern_config_lines = getConfigLines(flow_pattern_config_filename)
	#print "Flow Pattern config file contents:"
	#for line in flow_pattern_config_lines: print "\t%s"%line
	DSS_map_lines = getConfigLines(DSS_map_filename)
	#print "DSS map config file contents:"
	#for line in DSS_map_lines: print "\t%s"%line

	# Extract the configured DSS file and pathname for each reservoir's daily
	# disaggregation pattern.
	for line in flow_pattern_config_lines:
		token = line.strip().split(',')
		if len(token) != 3:
			print "File %s line \n\t \"%s\"\nis not a valid ID for a flow pattern DSS record."%(flow_pattern_config_filename,line)
			continue
		if line.split(',')[0].strip().upper() == "SHASTA":
			shasta_pattern_DSS_file_name = line.split(',')[1].strip().strip('\\')
			shasta_pattern_path = line.split(',')[2].strip()
		if line.split(',')[0].strip().upper() == "WHISKEYTOWN":
			whiskeytown_pattern_DSS_file_name = line.split(',')[1].strip().strip('\\')
			whiskeytown_pattern_path = line.split(',')[2].strip()
		if line.split(',')[0].strip().upper() == "TRINITY-CLAIR ENGLE":
			trinity_pattern_DSS_file_name = line.split(',')[1].strip().strip('\\')
			trinity_pattern_path = line.split(',')[2].strip()

	# All three reservoirs require both a pattern DSS file and pathname before
	# boundary-condition generation can continue.
	if (len(shasta_pattern_DSS_file_name) == 0 or len(whiskeytown_pattern_DSS_file_name) == 0 or
		len(trinity_pattern_DSS_file_name) == 0 or len(shasta_pattern_path) == 0 or
		len(whiskeytown_pattern_path) == 0 or len(trinity_pattern_path) == 0):
		print "Error reading flow pattern configuration file\n\t%s"%(flow_pattern_config_filename)
		print "Pattern DSS file or path not found."
		return None

	# Resolve configured pattern files against the project workspace when they
	# are supplied as relative paths.
	if not os.path.isabs(shasta_pattern_DSS_file_name):
		shasta_pattern_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), shasta_pattern_DSS_file_name)
	if not os.path.isabs(whiskeytown_pattern_DSS_file_name):
		whiskeytown_pattern_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), whiskeytown_pattern_DSS_file_name)
	if not os.path.isabs(trinity_pattern_DSS_file_name):
		trinity_pattern_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), trinity_pattern_DSS_file_name)

	# Locate the Redding Airport air-temperature record previously written to
	# the DSS map; it will drive the Shasta tributary temperature regressions.
	met_DSS_file_name = ""
	airtemp_path_redding = ""
	with open(DSS_map_filename) as infile:
		for line in infile:
			if (line.split(',')[0].strip().upper() == "REDDING AIRPORT" and
				line.split(',')[1].strip().upper() == "AIR TEMPERATURE"):
				met_DSS_file_name = line.split(',')[2].strip().strip('\\')
				airtemp_path_redding = line.split(',')[3].strip()

	# Stop if the required meteorological DSS reference is absent from the map.
	if len(met_DSS_file_name) == 0 or len(airtemp_path_redding) == 0:
		print "Error reading Shasta air temperature data configuration from file\n\t%s"%(DSS_map_filename)
		print "Air temperature DSS file or path not found."
		return None
	if not os.path.isabs(met_DSS_file_name):
		met_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), met_DSS_file_name)

	# Collect all generated time series for final DSS output and separately
	# retain reservoir balance series for closure checking.
	tsm_list = []
	balance_list = []

	########################
	# Trinity-Clair Engle and Lewiston
	# data from CVP spreadsheet
	########################
	print "TS Location = %s"%(trinity_tsc_list[0].location.upper())

	# Initialize daily accumulated-depletion flow and monthly volume-balance
	# series used to account for Trinity Lake inflows, outflows, and storage.
	tsmath_acc_dep = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_acc_dep.setUnits("CFS")
	tsmath_acc_dep.setType("PER-AVER")
	tsmath_acc_dep.setTimeInterval("1DAY")
	tsmath_acc_dep.setWatershed("TRINITY RIVER")
	tsmath_acc_dep.setLocation("TRINITY LAKE")
	tsmath_acc_dep.setParameterPart("FLOW-ACC-DEP")
	tsmath_acc_dep.setVersion(BC_F_part)

	tsmath_bal_trnty = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1MONTH", "0M", 0.0)
	tsmath_bal_trnty.setUnits("AC-FT")
	tsmath_bal_trnty.setType("PER-CUM")
	tsmath_bal_trnty.setTimeInterval("1MONTH")
	tsmath_bal_trnty.setWatershed("TRINITY RIVER")
	tsmath_bal_trnty.setLocation("TRINITY LAKE")
	tsmath_bal_trnty.setParameterPart("VOLUME-BALANCE")
	tsmath_bal_trnty.setVersion(BC_F_part)

	# Interpret each imported Trinity series according to its parameter and
	# construct the corresponding boundary-condition and balance components.
	for ts in trinity_tsc_list:
		print "\tTS Parameter = %s"%(ts.parameter.upper())
		tsm = tsmath(ts)
		tsm.setWatershed("TRINITY RIVER")
		tsm.setLocation("TRINITY LAKE")

		# Add monthly inflow to the volume balance and disaggregate it to daily
		# flow using the configured Trinity pattern.
		if ts.parameter.upper() == "INFLOW":
			tsmath_flow_monthly = tsm
			tsm_list.append(tsmath_flow_monthly)
			tsmath_bal_trnty = tsmath_bal_trnty.add(tsmath_flow_monthly)
			print "reading Trinity pattern from file: " + trinity_pattern_DSS_file_name
			print "\tDSS path:" + trinity_pattern_path
			ts_read = hec.heclib.dss.HecTimeSeries()
			ts_read.setDSSFileName(trinity_pattern_DSS_file_name)
			tsc_pattern = tscont()
			tsc_pattern.fullName = trinity_pattern_path
			status = ts_read.read(tsc_pattern, False)
			if status < 0:
				print "Failed to read pattern time series %s \n\tfrom DSS file %s"%(tsc_pattern.fullName, trinity_dss_file_name)
				ts_read.done()
				continue
			tsmath_pattern = tsmath(tsc_pattern)
			ts_read.done()

			tsmath_trinity_inflow_daily = CVP.weight_transform_monthly_to_daily(
				tsmath_flow_monthly, tsmath_pattern, start_day_count=days_in_first_month)
			tsmath_trinity_inflow_daily.setPathname(ts.fullName)
			tsmath_trinity_inflow_daily.setWatershed("TRINITY RIVER")
			tsmath_trinity_inflow_daily.setLocation("TRINITY LAKE")
			tsmath_trinity_inflow_daily.setTimeInterval("1DAY")
			tsmath_trinity_inflow_daily.setParameterPart("FLOW-IN")
			tsmath_trinity_inflow_daily.setVersion(BC_F_part)
			tsm_list.append(tsmath_trinity_inflow_daily)

		# Convert monthly storage to instantaneous storage and derive successive
		# storage changes for inclusion in the monthly reservoir balance.
		elif "STORAGE" in ts.parameter.upper():
			tsmath_storage_monthly = tsm
			tsmath_storage_monthly.setParameterPart("STORAGE")
			tsmath_storage_monthly.setType("INST-CUM")
			tsm_storage_change = tsmath_storage_monthly.successiveDifferences()
			tsmath_storage_monthly.setType("INST-VAL")
			tsm_list.append(tsmath_storage_monthly)
			tsm_storage_change.setWatershed("TRINITY RIVER")
			tsm_storage_change.setLocation("TRINITY LAKE")
			tsm_storage_change.setParameterPart("STORAGE-CHANGE")
			tsm_list.append(tsm_storage_change)
			tsmath_bal_trnty = tsmath_bal_trnty.subtract(tsm_storage_change)

		# Treat evaporation as a monthly volume loss and convert it to a daily
		# flow contribution for the accumulated-depletion term.
		elif "EVAP" in ts.parameter.upper():
			tsmath_evap_monthly = tsm
			tsmath_evap_monthly.setParameterPart("VOLUME-EST EVAPORATION")
			tsm_list.append(tsmath_evap_monthly)
			tsmath_bal_trnty = tsmath_bal_trnty.subtract(tsmath_evap_monthly)
			tsmath_acc_dep = tsmath_acc_dep.subtract(
				CVP.uniform_transform_monthly_to_daily(
				tsmath_evap_monthly, start_day_count=days_in_first_month))
			tsm_list.append(CVP.uniform_transform_monthly_to_daily(
				tsmath_evap_monthly, start_day_count=days_in_first_month))

		# Convert total monthly Trinity releases to an hourly release-flow
		# boundary condition while retaining the monthly volume record.
		elif ts.parameter.upper() == "TOTAL RELEASE":
			tsmath_trinity_release_monthly = tsm
			tsmath_trinity_release_monthly.setParameterPart("VOLUME-RELEASE")
			tsm_list.append(tsmath_trinity_release_monthly)
			tsmath_trinity_release = CVP.uniform_transform_monthly_to_hourly(
				tsmath_trinity_release_monthly, start_day_count=days_in_first_month)
			tsmath_trinity_release.setWatershed("TRINITY RIVER")
			tsmath_trinity_release.setLocation("TRINITY LAKE")
			tsmath_trinity_release.setParameterPart("FLOW-RELEASE")
			tsmath_trinity_release.setTimeInterval("1HOUR")
			tsmath_trinity_release.setVersion(BC_F_part)
			tsm_list.append(tsmath_trinity_release)

		# Retain the CFS Lewiston river-release record with reservoir-specific
		# location and parameter metadata.
		elif "RIVER REL" in ts.parameter.upper() and "CFS" in ts.parameter.upper():
			tsmath_lewiston_release_flow_monthly = tsm
			tsmath_lewiston_release_flow_monthly.setLocation("LEWISTON RESERVOIR")
			tsmath_lewiston_release_flow_monthly.setParameterPart("FLOW-RIVER RELEASE")
			# tsm_list.append(tsmath_lewiston_release_flow_monthly)

		# Subtract the monthly Lewiston release volume from the Trinity balance
		# and create its corresponding daily flow series.
		elif "RIVER REL" in ts.parameter.upper() and "TAF" in ts.parameter.upper():
			tsmath_lewiston_release_monthly = tsm
			tsmath_lewiston_release_monthly.setLocation("LEWISTON RESERVOIR")
			tsmath_lewiston_release_monthly.setParameterPart("VOLUME-RIVER RELEASE")
			tsm_list.append(tsmath_lewiston_release_monthly)
			tsmath_bal_trnty = tsmath_bal_trnty.subtract(tsmath_lewiston_release_monthly)
			tsmath_lewiston_release = CVP.uniform_transform_monthly_to_daily(
				tsmath_lewiston_release_monthly, start_day_count=days_in_first_month)
			tsmath_lewiston_release.setWatershed("TRINITY RIVER")
			tsmath_lewiston_release.setLocation("LEWISTON RESERVOIR")
			tsmath_lewiston_release.setParameterPart("FLOW-RIVER RELEASE")
			tsmath_lewiston_release.setTimeInterval("1DAY")
			tsmath_lewiston_release.setVersion(BC_F_part)
			tsm_list.append(tsmath_lewiston_release)

		# Treat Carr Powerhouse flow as a Clear Creek diversion from the Trinity
		# balance and generate the corresponding hourly release series.
		elif ts.parameter.upper() == "CARR PP":
			tsmath_carr_release_monthly = tsm
			tsmath_carr_release_monthly.setWatershed("TRINITY RIVER")
			tsmath_carr_release_monthly.setLocation("LEWISTON RESERVOIR")
			tsmath_carr_release_monthly.setParameterPart("VOLUME-CLEAR CREEK DIVERSION")
			tsm_list.append(tsmath_carr_release_monthly)
			tsmath_bal_trnty = tsmath_bal_trnty.subtract(tsmath_carr_release_monthly)
			tsmath_carr_release = CVP.uniform_transform_monthly_to_hourly(
				tsm, start_day_count=days_in_first_month)
			tsmath_carr_release.setWatershed("CLEAR CREEK")
			tsmath_carr_release.setLocation("CARR POWERHOUSE")
			tsmath_carr_release.setParameterPart("FLOW-RELEASE")
			tsmath_carr_release.setTimeInterval("1HOUR")
			tsmath_carr_release.setVersion(BC_F_part)
			tsm_list.append(tsmath_carr_release)
		else:
			tsm_list.append(tsm)

	# Trinity storage changes due to:
	#	In:
	#		Trinity inflow : tsmath_trinity_inflow_daily
	#	Out:
	#		Trinity dam releases: tsmath_release_daily
	#		Net evaporation, leakage, other: tsmath_acc_dep

	# Construct a daily Trinity storage trajectory, using monthly storage values
	# as fixed checkpoints and the daily water balance between those checkpoints.
	tsmath_storage_daily = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_storage_daily.setUnits("AC-FT")
	tsmath_storage_daily.setType("INST-VAL")
	tsmath_storage_daily.setTimeInterval("1DAY")
	tsmath_storage_daily.setWatershed("TRINITY RIVER")
	tsmath_storage_daily.setLocation("TRINITY LAKE")
	tsmath_storage_daily.setParameterPart("STORAGE-CVP")
	tsmath_storage_daily.setVersion(BC_F_part)
	tsmath_storage_daily.getContainer().values[0] = tsmath_storage_monthly.getContainer().values[0]
	tsmath_release_daily = CVP.uniform_transform_monthly_to_daily(
		tsmath_trinity_release_monthly, start_day_count=days_in_first_month)

	j = 1
	search_time = HecTime()

	# At each monthly checkpoint, reset to the imported storage value; otherwise
	# integrate the daily inflow, release, and accumulated-depletion balance.
	for i in range(1, len(tsmath_storage_daily.getContainer().values)):
		if tsmath_storage_daily.getContainer().times[i] >= tsmath_storage_monthly.getContainer().times[j]:
			tsmath_storage_daily.getContainer().values[i] = tsmath_storage_monthly.getContainer().values[j]
			j += 1
		else:
			search_time.set(tsmath_storage_daily.getContainer().times[i])
			tsmath_storage_daily.getContainer().values[i] = (
				tsmath_storage_daily.getContainer().values[i-1] + 1.98347*(
				tsmath_trinity_inflow_daily.getContainer().getValue(search_time)
				- tsmath_release_daily.getContainer().getValue(search_time)
				+ tsmath_acc_dep.getContainer().getValue(search_time)))

	# Retain the calculated daily storage, accumulated-depletion, and monthly
	# balance series for final output and closure checking.
	tsm_list.append(tsmath_storage_daily)
	tsm_list.append(tsmath_acc_dep)
	tsm_list.append(tsmath_bal_trnty)
	balance_list.append(tsmath_bal_trnty)

	########################
	# Disaggregate Trinity Lake Tributary In Flows
	########################

	# Table of tributary weights by month
	tributary_weights = {
		"EF TRINITY":(0.200999652, 0.201001041, 0.200998676, 0.200998663, 0.201001043, 0.200998535, 0.200996912, 0.201009666, 0.200970155, 0.200978909, 0.201003315, 0.201000536),
		"STUART FORK":(0.119998966, 0.119998249, 0.120001072, 0.120000814, 0.119998453, 0.119997145, 0.120002941, 0.120069866, 0.120091685, 0.120009144, 0.119987691, 0.120004223),
		"SWIFT CR":(0.114000898, 0.114002108, 0.114000033, 0.114001063, 0.113999875, 0.114002195, 0.114015586, 0.113944013, 0.114068202, 0.114014717, 0.113997178, 0.113996712),
		"TRINITY RIVER":(0.565000485, 0.564998603, 0.565000218, 0.56499946, 0.56500063, 0.565002124, 0.564984561, 0.564976455, 0.564869961, 0.564997226, 0.565011814, 0.56499853)}

	# Split total Trinity inflow among the four modeled tributaries using the
	# month-specific fractions above and retain each series by location name.
	names_flows = {}
	for tsm in CVP.split_time_series_monthly(tsmath_trinity_inflow_daily, tributary_weights, "FLOW-IN"):
		tsm.setVersion(BC_F_part)
		tsm_list.append(tsm)
		names_flows[tsm.getContainer().location] = tsm

	########################
	# Estimate Trinity Tributary Temperatures
	########################

	# Locate the Lewiston air-temperature DSS record used as the meteorological
	# predictor for Trinity tributary temperature regressions.
	met_DSS_file_name = ""
	airtemp_path_lewiston = ""

	for line in DSS_map_lines:
		# print line
		if (line.split(',')[0].strip().upper() == "LEWISTON RES" and
			line.split(',')[1].strip().upper() == "AIR TEMPERATURE"):
			met_DSS_file_name = line.split(',')[2].strip().strip('\\')
			airtemp_path_lewiston = line.split(',')[3].strip()
			break

	# Require a valid DSS file/path mapping before attempting the regressions.
	if len(met_DSS_file_name) == 0 or len(airtemp_path_lewiston) == 0:
		print "Error reading Trinity air temperature data configuration from file\n\t%s"%(DSS_map_filename)
		print "Air temperature DSS file or path not found."
		return None
	if not os.path.isabs(met_DSS_file_name):
		met_DSS_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), met_DSS_file_name)

	#River, Intercept (deg C), Flow Coef (cfs), Air Temp Coef (deg C), RMS Error (deg C)
	tributary_temp_regression_coefficients = {
		"EF TRINITY": (2.204979, -0.00208361, 0.65876114, 2.117),
		"STUART FORK": (1.2766113, -0.00304511, 0.60274446, 1.9729857),
		"SWIFT CR": (1.2773657, -0.00356459,  0.6329333, 2.0825596),
		"TRINITY RIVER": (1.968627, -0.00075939, 0.6476875, 2.102819)}

	# Read the Lewiston air-temperature series from DSS.
	ts_read = hec.heclib.dss.HecTimeSeries()
	ts_read.setDSSFileName(met_DSS_file_name)
	tsc_airtemp = tscont()
	tsc_airtemp.fullName = airtemp_path_lewiston
	status = ts_read.read(tsc_airtemp, False)
	if status < 0:
		print "Failed to read pattern time series %s \n\tfrom DSS file %s"%(tsc_pattern.fullName, whiskeytown_pattern_DSS_file_name)
		ts_read.done()
	tsmath_airtemp = tsmath(tsc_airtemp)
	ts_read.done()

	# Estimate water temperature independently for each disaggregated Trinity
	# tributary using its flow series and the common Lewiston air temperature.
	for key in tributary_temp_regression_coefficients.keys():
		tsm = CVP.evaluate_temp_regression(names_flows[key], tsmath_airtemp, tributary_temp_regression_coefficients[key])
		tsm.setVersion(BC_F_part)
		tsm_list.append(tsm)

	########################
	# Whiskeytown and Clear Creek
	# data from CVP spreadsheet
	########################
	print "TS Location = %s"%(whiskeytown_tsc_list[0].location.upper())

	# Initialize the daily accumulated-depletion term and monthly water-balance
	# series for Whiskeytown Lake and the Clear Creek system.
	tsmath_acc_dep = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_acc_dep.setUnits("CFS")
	tsmath_acc_dep.setType("PER-AVER")
	tsmath_acc_dep.setTimeInterval("1DAY")
	tsmath_acc_dep.setWatershed("CLEAR CREEK")
	tsmath_acc_dep.setLocation("WHISKEYTOWN LAKE")
	tsmath_acc_dep.setParameterPart("FLOW-ACC-DEP")
	tsmath_acc_dep.setVersion(BC_F_part)

	tsmath_bal_whsky = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1MONTH", "0M", 0.0)
	tsmath_bal_whsky.setUnits("AC-FT")
	tsmath_bal_whsky.setType("PER-CUM")
	tsmath_bal_whsky.setTimeInterval("1MONTH")
	tsmath_bal_whsky.setWatershed("CLEAR CREEK")
	tsmath_bal_whsky.setLocation("WHISKEYTOWN LAKE")
	tsmath_bal_whsky.setParameterPart("VOLUME-BALANCE")
	tsmath_bal_whsky.setVersion(BC_F_part)

	# Carr Powerhouse releases enter the Whiskeytown/Clear Creek system and
	# therefore contribute positively to the Whiskeytown volume balance.
	tsmath_bal_whsky = tsmath_bal_whsky.add(tsmath_carr_release_monthly)

	# Interpret each imported Whiskeytown series and construct the corresponding
	# boundary-condition and reservoir-balance components.
	for ts in whiskeytown_tsc_list:
		print "\tTS Parameter = %s"%(ts.parameter.upper())
		tsm = tsmath(ts)
		tsm.setWatershed("CLEAR CREEK")
		tsm.setLocation("WHISKEYTOWN LAKE")

		# Add monthly inflow to the Whiskeytown balance and use the configured
		# daily pattern to disaggregate it to daily flow.
		if ts.parameter.upper() == "INFLOW":
			tsmath_flow_monthly = tsm
			tsm_list.append(tsmath_flow_monthly)
			tsmath_bal_whsky = tsmath_bal_whsky.add(tsmath_flow_monthly)
			print "reading pattern from file: " + whiskeytown_pattern_DSS_file_name
			print "\t" + whiskeytown_pattern_path
			ts_read = hec.heclib.dss.HecTimeSeries()
			ts_read.setDSSFileName(whiskeytown_pattern_DSS_file_name)
			tsc_pattern = tscont()
			tsc_pattern.fullName = whiskeytown_pattern_path
			status = ts_read.read(tsc_pattern, False)
			if status < 0:
				print "Failed to read pattern time series %s \n\tfrom DSS file %s"%(tsc_pattern.fullName, whiskeytown_pattern_DSS_file_name)
				ts_read.done()
				continue
			tsmath_pattern = tsmath(tsc_pattern)
			ts_read.done()

			tsmath_weighted = CVP.weight_transform_monthly_to_daily(
				tsmath_flow_monthly, tsmath_pattern, start_day_count=days_in_first_month)
			tsmath_weighted.setPathname(ts.fullName)
			tsmath_weighted.setTimeInterval("1DAY")
			tsmath_weighted.setParameterPart("FLOW-IN")
			tsmath_weighted.setVersion(BC_F_part)
			tsm_list.append(tsmath_weighted)

		# Convert monthly storage to instantaneous storage and derive successive
		# storage changes for the monthly volume balance.
		elif "STORAGE" in ts.parameter.upper():
			tsmath_storage_monthly = tsm
			tsmath_storage_monthly.setParameterPart("STORAGE")
			tsmath_storage_monthly.setType("INST-CUM")
			tsm_storage_change = tsmath_storage_monthly.successiveDifferences()
			tsmath_storage_monthly.setType("INST-VAL")
			tsm_list.append(tsmath_storage_monthly)
			tsm_storage_change.setWatershed("CLEAR CREEK")
			tsm_storage_change.setLocation("WHISKEYTOWN LAKE")
			tsm_storage_change.setParameterPart("STORAGE-CHANGE")
			tsm_list.append(tsm_storage_change)
			tsmath_bal_whsky = tsmath_bal_whsky.subtract(tsm_storage_change)

		# Treat Spring Creek as an outflow from the Whiskeytown balance and
		# convert its monthly volume series to an hourly flow boundary condition.
		elif "SPRING CR" in ts.parameter.upper():
			tsmath_sp_cr_monthly = tsm
			tsm_list.append(tsmath_sp_cr_monthly)
			tsmath_bal_whsky = tsmath_bal_whsky.subtract(tsmath_sp_cr_monthly)
			tsmath_sp_cr = CVP.uniform_transform_monthly_to_hourly(
				tsm, start_day_count=days_in_first_month)
			tsmath_sp_cr.setPathname(ts.fullName)
			tsmath_sp_cr.setLocation("SPRING CREEK")
			tsmath_sp_cr.setTimeInterval("1HOUR")
			tsmath_sp_cr.setParameterPart("FLOW-PP")
			tsmath_sp_cr.setVersion(BC_F_part)
			tsm_list.append(tsmath_sp_cr)

		# Subtract evaporation from the monthly reservoir balance and represent
		# the same loss as a daily accumulated-depletion flow.
		elif "EVAP" in ts.parameter.upper():
			tsmath_evap_monthly = tsm
			tsmath_evap_monthly.setParameterPart("VOLUME-EST EVAPORATION")
			tsm_list.append(tsmath_evap_monthly)
			tsmath_bal_whsky = tsmath_bal_whsky.subtract(tsmath_evap_monthly)
			tsmath_acc_dep = tsmath_acc_dep.subtract(
				CVP.uniform_transform_monthly_to_daily(
				tsmath_evap_monthly, start_day_count=days_in_first_month))

		# Process Clear Creek release volume as a Whiskeytown Dam outflow and
		# generate the corresponding hourly release-flow series.
		elif "CLEAR CREEK" in ts.parameter.upper() and "TAF" in ts.parameter.upper():
			tsmath_release_monthly = tsm
			tsmath_release_monthly.setLocation("WHISKEYTOWN DAM")
			tsm_list.append(tsmath_release_monthly)
			tsmath_bal_whsky = tsmath_bal_whsky.subtract(tsmath_release_monthly)
			tsmath_release = CVP.uniform_transform_monthly_to_hourly(
				tsmath_release_monthly, start_day_count=days_in_first_month)
			tsmath_release.setPathname(tsmath_release_monthly.getContainer().fullName)
			tsmath_release.setLocation("WHISKEYTOWN DAM")
			tsmath_release.setTimeInterval("1HOUR")
			tsmath_release.setParameterPart("FLOW-RELEASE")
			tsmath_release.setVersion(BC_F_part)
			tsm_list.append(tsmath_release)
		else:
			tsm_list.append(tsm)

	# Whiskeytown storage changes due to:
	#	In:
	#		Clear Creek inflow: tsmath_weighted
	#		Carr Powerhouse releases: tsmath_carr_release_daily
	#	Out:
	#		Whiskeytown dam releases: tsmath_release_daily
	#		Spring Creek Tunnel releases: tsmath_sp_cr_daily
	#		Net evaporation, leakage, other: tsmath_acc_dep

	# Construct daily Whiskeytown storage using imported monthly storage as
	# checkpoints and the daily water balance between those checkpoints.
	tsmath_storage_daily = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_storage_daily.setUnits("AC-FT")
	tsmath_storage_daily.setType("INST-VAL")
	tsmath_storage_daily.setTimeInterval("1DAY")
	tsmath_storage_daily.setWatershed("CLEAR CREEK")
	tsmath_storage_daily.setLocation("WHISKEYTOWN LAKE")
	tsmath_storage_daily.setParameterPart("STORAGE-CVP")
	tsmath_storage_daily.setVersion(BC_F_part)
	tsmath_storage_daily.getContainer().values[0] = tsmath_storage_monthly.getContainer().values[0]

	# Convert each monthly Whiskeytown balance-flow component to daily values
	# for use in the storage integration below.
	tsmath_release_daily = CVP.uniform_transform_monthly_to_daily(
		tsmath_release_monthly, start_day_count=days_in_first_month)
	tsmath_sp_cr_daily = CVP.uniform_transform_monthly_to_daily(
		tsmath_sp_cr_monthly, start_day_count=days_in_first_month)
	tsmath_carr_release_daily = CVP.uniform_transform_monthly_to_daily(
		tsmath_carr_release_monthly, start_day_count=days_in_first_month)

	j = 1
	search_time = HecTime()

	# Reset storage to the imported monthly value at each checkpoint; on other
	# days, integrate inflows, releases, diversions, and accumulated depletion.
	for i in range(1, len(tsmath_storage_daily.getContainer().values)):
		if tsmath_storage_daily.getContainer().times[i] >= tsmath_storage_monthly.getContainer().times[j]:
			tsmath_storage_daily.getContainer().values[i] = tsmath_storage_monthly.getContainer().values[j]
			j += 1
		else:
			search_time.set(tsmath_storage_daily.getContainer().times[i])
			tsmath_storage_daily.getContainer().values[i] = (
				tsmath_storage_daily.getContainer().values[i-1] + 1.98347*(
				tsmath_weighted.getContainer().getValue(search_time)
				+ tsmath_carr_release_daily.getContainer().getValue(search_time)
				- tsmath_release_daily.getContainer().getValue(search_time)
				- tsmath_sp_cr_daily.getContainer().getValue(search_time)
				+ tsmath_acc_dep.getContainer().getValue(search_time)))

	# Retain the daily storage, accumulated-depletion, and monthly balance
	# records for final DSS output and closure checking.
	tsm_list.append(tsmath_storage_daily)

	tsm_list.append(tsmath_acc_dep)
	tsm_list.append(tsmath_bal_whsky)
	balance_list.append(tsmath_bal_whsky)


	########################
	# Shasta/Keswick & main-stem Sacramento data from CVP spreadsheet
	########################
	print "TS Location = %s"%(shasta_tsc_list[0].location.upper())

	# Initialize the Shasta daily accumulated-depletion term and monthly
	# reservoir volume-balance series.
	tsmath_acc_dep = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_acc_dep.setUnits("CFS")
	tsmath_acc_dep.setType("PER-AVER")
	tsmath_acc_dep.setTimeInterval("1DAY")
	tsmath_acc_dep.setWatershed("SACRAMENTO RIVER")
	tsmath_acc_dep.setLocation("SHASTA LAKE")
	tsmath_acc_dep.setParameterPart("FLOW-ACC-DEP")
	tsmath_acc_dep.setVersion(BC_F_part)

	tsmath_bal_shasta = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1MONTH", "0M", 0.0)
	tsmath_bal_shasta.setUnits("AC-FT")
	tsmath_bal_shasta.setType("PER-CUM")
	tsmath_bal_shasta.setTimeInterval("1MONTH")
	tsmath_bal_shasta.setWatershed("SACRAMENTO RIVER")
	tsmath_bal_shasta.setLocation("SHASTA LAKE")
	tsmath_bal_shasta.setParameterPart("VOLUME-BALANCE")
	tsmath_bal_shasta.setVersion(BC_F_part)

	# Interpret each Shasta spreadsheet series according to its parameter and
	# build the associated boundary-condition and reservoir-balance records.
	for ts in shasta_tsc_list:
		print "\tTS Parameter = %s"%(ts.parameter.upper())
		tsm = tsmath(ts)
		tsm.setWatershed("SACRAMENTO RIVER")
		tsm.setLocation("SHASTA LAKE")

		# Add monthly Shasta inflow to the volume balance and disaggregate it
		# using the configured daily flow pattern.
		if "INFLOW" in ts.parameter.upper():
			tsmath_flow_monthly = tsm
			tsm_list.append(tsmath_flow_monthly)
			tsmath_bal_shasta = tsmath_bal_shasta.add(tsmath_flow_monthly)
			print "\treading pattern from file: " + shasta_pattern_DSS_file_name
			print "\t\t" + shasta_pattern_path
			ts_read = hec.heclib.dss.HecTimeSeries()
			ts_read.setDSSFileName(shasta_pattern_DSS_file_name)
			tsc_pattern = tscont()
			tsc_pattern.fullName = shasta_pattern_path
			status = ts_read.read(tsc_pattern, False)
			if status < 0:
				print "Failed to read pattern time series %s \n\tfrom DSS file %s"%(tsc_pattern.fullName, shasta_pattern_DSS_file_name)
				ts_read.done()
				continue
			tsmath_pattern = tsmath(tsc_pattern)
			ts_read.done()

			tsmath_weighted = CVP.weight_transform_monthly_to_daily(
				tsmath_flow_monthly, tsmath_pattern, start_day_count=days_in_first_month)
			tsmath_weighted.setPathname(ts.fullName)
			tsmath_weighted.setTimeInterval("1DAY")
			tsmath_weighted.setParameterPart("FLOW-IN")
			tsmath_weighted.setVersion(BC_F_part)
			tsm_list.append(tsmath_weighted)

		# Convert imported monthly Shasta storage to instantaneous storage and
		# derive the month-to-month change for the volume balance.
		elif "STORAGE" in ts.parameter.upper():
			tsmath_storage_monthly = tsm
			tsmath_storage_monthly.setParameterPart("STORAGE")
			tsmath_storage_monthly.setType("INST-CUM")
			tsm_storage_change = tsmath_storage_monthly.successiveDifferences()
			tsmath_storage_monthly.setType("INST-VAL")
			tsm_list.append(tsmath_storage_monthly)
			tsm_storage_change.setWatershed("SACRAMENTO RIVER")
			tsm_storage_change.setLocation("SHASTA LAKE")
			tsm_storage_change.setParameterPart("STORAGE-CHANGE")
			tsm_list.append(tsm_storage_change)
			tsmath_bal_shasta = tsmath_bal_shasta.subtract(tsm_storage_change)

		# Subtract evaporation from the monthly Shasta balance and represent the
		# loss as a daily accumulated-depletion flow.
		elif "EVAP" in ts.parameter.upper():
			tsmath_evap_monthly = tsm
			tsmath_evap_monthly.setParameterPart("VOLUME-EST EVAPORATION")
			tsm_list.append(tsmath_evap_monthly)
			tsmath_bal_shasta = tsmath_bal_shasta.subtract(tsmath_evap_monthly)
			tsmath_acc_dep = tsmath_acc_dep.subtract(
				CVP.uniform_transform_monthly_to_daily(
					tsmath_evap_monthly, start_day_count=days_in_first_month))

		# Convert total Shasta release volume to an hourly release-flow series
		# and subtract the monthly release from the reservoir balance.
		elif ts.parameter.upper() == "TOTAL SHASTA RELEASE":
			tsmath_release_monthly = tsm
			tsm_list.append(tsmath_release_monthly)
			tsmath_bal_shasta = tsmath_bal_shasta.subtract(tsmath_release_monthly)
			tsmath_release = CVP.uniform_transform_monthly_to_hourly(
				tsmath_release_monthly, start_day_count=days_in_first_month)
			tsmath_release.setPathname(ts.fullName)
			tsmath_release.setTimeInterval("1HOUR")
			tsmath_release.setParameterPart("FLOW-RELEASE")
			tsmath_release.setVersion(BC_F_part)
			tsm_list.append(tsmath_release)

		# Preserve the Keswick flow series and create the corresponding hourly
		# boundary-condition representation without modifying the Shasta balance.
		elif ts.parameter.upper() == "FLOW-KESWICK-CFS":
			tsmath_release_monthly = tsm
			tsm_list.append(tsmath_release_monthly)
			tsmath_release = CVP.uniform_transform_monthly_to_hourly(
				tsmath_release_monthly, start_day_count=days_in_first_month)
			tsmath_release.setPathname(ts.fullName)
			tsmath_release.setTimeInterval("1HOUR")
			tsmath_release.setParameterPart("FLOW-RELEASE-KESWICK-CFS")
			tsmath_release.setVersion(BC_F_part)
			tsm_list.append(tsmath_release)
		else:
			tsm_list.append(tsm)

	# Shasta storage changes due to:
	#	In:
	#		Shasta inflow: tsmath_weighted
	#		Carr Powerhouse releases: tsmath_carr_release_daily
	#	Out:
	#		Shasta dam releases: tsmath_release_daily
	#		Net evaporation, leakage, other: tsmath_acc_dep

	# Construct daily Shasta storage using the imported monthly storage values
	# as checkpoints and the intervening daily reservoir water balance.
	tsmath_storage_daily = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_storage_daily.setUnits("AC-FT")
	tsmath_storage_daily.setType("INST-VAL")
	tsmath_storage_daily.setTimeInterval("1DAY")
	tsmath_storage_daily.setWatershed("SACRAMENTO RIVER")
	tsmath_storage_daily.setLocation("SHASTA LAKE")
	tsmath_storage_daily.setParameterPart("STORAGE-CVP")
	tsmath_storage_daily.setVersion(BC_F_part)
	tsmath_storage_daily.getContainer().values[0] = tsmath_storage_monthly.getContainer().values[0]
	tsmath_release_daily = CVP.uniform_transform_monthly_to_daily(
		tsmath_release_monthly, start_day_count=days_in_first_month)

	j = 1
	search_time = HecTime()

	# Reset storage to the imported monthly value at each checkpoint; otherwise
	# integrate the daily inflows, releases, and accumulated-depletion terms.
	for i in range(1, len(tsmath_storage_daily.getContainer().values)):
		if tsmath_storage_daily.getContainer().times[i] >= tsmath_storage_monthly.getContainer().times[j]:
			tsmath_storage_daily.getContainer().values[i] = tsmath_storage_monthly.getContainer().values[j]
			j += 1
		else:
			search_time.set(tsmath_storage_daily.getContainer().times[i])
			tsmath_storage_daily.getContainer().values[i] = (
				tsmath_storage_daily.getContainer().values[i-1] + 1.98347*(
				tsmath_weighted.getContainer().getValue(search_time)
				+ tsmath_carr_release_daily.getContainer().getValue(search_time)
				- tsmath_release_daily.getContainer().getValue(search_time)
				- tsmath_sp_cr_daily.getContainer().getValue(search_time)
				+ tsmath_acc_dep.getContainer().getValue(search_time)))

	# Retain the calculated Shasta storage and balance terms for final DSS
	# output and subsequent reservoir-balance closure checks.
	tsm_list.append(tsmath_storage_daily)


	tsm_list.append(tsmath_acc_dep)
	tsm_list.append(tsmath_bal_shasta)
	balance_list.append(tsmath_bal_shasta)

	########################
	# Disaggregate Shasta Tributary In Flows
	########################

	# Define the month-specific fractions used to partition total Shasta inflow
	# among the modeled Sacramento, McCloud, Sulanharas, and Pit tributaries.
	tributary_weights = {
		"Shasta-Sac-in":(0.212770745, 0.224327192, 0.221179858, 0.231031865, 0.22998096, 0.174508497, 0.096498474, 0.074162081, 0.066134982, 0.085930713, 0.110001981, 0.208573952),
		"Shasta-McCloud-in":(0.138567582, 0.157547951, 0.139190927, 0.129798785, 0.107066929, 0.097430013, 0.099133208, 0.094616182, 0.097972639, 0.111942455, 0.109353393, 0.151801944),
		"Shasta-Sulanharas-in":(0.037029058, 0.042679679, 0.040961806, 0.039603204, 0.037053932, 0.024035906, 0.01008518, 0.006934946, 0.006026154, 0.009676144, 0.013545787, 0.038994156),
		"Shasta-Pit-in":(0.611632586, 0.575445235, 0.598667383, 0.599566102, 0.625898182, 0.704025567, 0.794283211, 0.824286819, 0.82986623, 0.792450666, 0.767098904, 0.600629926)}

	# Split the total daily Shasta inflow according to the monthly fractions and
	# retain each generated tributary series by location for temperature modeling.
	names_flows = {}
	for tsm in CVP.split_time_series_monthly(tsmath_weighted, tributary_weights, "FLOW-IN"):
		tsm.setVersion(BC_F_part)
		tsm_list.append(tsm)
		names_flows[tsm.getContainer().location] = tsm

	########################
	# Estimate Shasta Tributary Temperatures
	########################

	# Regression tuples contain intercept, flow coefficient, air-temperature
	# coefficient, and RMSE, with temperature quantities expressed in degrees C.
	#River, Intercept (deg C), Flow Coef (cfs), Air Temp Coef (deg C), RMSE Error (deg C)
	tributary_temp_regression_coefficients = {
		"Shasta-Sac-in": (2.415046492, -0.000987216, 0.58256543, 1.857597335),
		"Shasta-Pit-in": (4.029028312, -0.000234506,0.510349367, 2.123152431),
		"Shasta-McCloud-in": (2.221323686, 5.18E-05, 0.463827771, 1.438501982)}

	# Read the Redding air-temperature series previously identified in the DSS
	# map for use as the meteorological predictor in the regressions.
	ts_read = hec.heclib.dss.HecTimeSeries()
	ts_read.setDSSFileName(met_DSS_file_name)
	tsc_airtemp = tscont()
	tsc_airtemp.fullName = airtemp_path_redding
	status = ts_read.read(tsc_airtemp, False)
	if status < 0:
		print "Failed to read pattern time series %s \n\tfrom DSS file %s"%(tsc_pattern.fullName, whiskeytown_pattern_DSS_file_name)
		ts_read.done()
	tsmath_airtemp = tsmath(tsc_airtemp)
	ts_read.done()

	# Estimate water temperature for each Shasta tributary having a configured
	# regression relationship and retain the resulting time series for output.
	for key in tributary_temp_regression_coefficients.keys():
		tsm = CVP.evaluate_temp_regression(names_flows[key], tsmath_airtemp, tributary_temp_regression_coefficients[key])
		tsm.setVersion(BC_F_part)
		tsm_list.append(tsm)

	########################
	# Get flows and temperatures for downstream tributaries
	# from monthly average data sets
	########################

	# Read each configured downstream-tributary monthly average, shift the
	# repeating monthly sequence to the forecast period, and convert it to daily.
	tributary_config_filename = os.path.join(Project.getCurrentProject().getWorkspacePath(), r"forecast\config\tributary_averages.config")
	trib_DSS_files = {}
	for line in getConfigLines(tributary_config_filename):
		token = line.split(',')
		dss_file_name = token[-2].strip()
		if not os.path.isabs(dss_file_name):
			dss_file_name = os.path.join(Project.getCurrentProject().getWorkspacePath(), dss_file_name)
		ts_read = hec.heclib.dss.HecTimeSeries()
		ts_read.setDSSFileName(dss_file_name)
		tsc_avg = tscont()
		tsc_avg.fullName = token[-1].strip()
		status = ts_read.read(tsc_avg, False)
		if status < 0:
			print "Failed to read temperature time series %s \n\tfrom DSS file %s"%(tsc_avg.fullName, dss_file_name)
			ts_read.done()
			continue
		tsmath_avg = tsmath(tsc_avg)
		tsmath_shift = shift_monthly_averages(tsmath_avg, start_time, end_time)
		shift_path = token[-1].strip().split('/')
		shift_path[6] = BC_F_part
		tsmath_shift.getContainer().fullName = '/'.join(shift_path)
		tsm_list.append(CVP.uniform_transform_monthly_to_daily(
			tsmath_shift, start_day_count=days_in_first_month))

	# Close any tributary DSS handles retained in the file dictionary.
	for fname in trib_DSS_files:
		trib_DSS_files[fname].done()

	########################
	# Check balances
	########################

	# Identify reservoir/month combinations where the calculated volume-balance
	# residual exceeds the allowed magnitude of 1,000 acre-feet.
	msg = "Volume balance failed to close within 1,000 AF at these locations and times: \n"
	exceed_list = []
	for tsm_bal in balance_list:
		balance_err = False
		err_type_neg = False
		if tsm_bal.max() > 1000:
			balance_err = True
		if tsm_bal.min() < -1000:
			balance_err = True
			err_type_neg = True
		if balance_err:
			if err_type_neg:
				exceed_list.append((tsm_bal.getContainer().location, tsm_bal.minDate()))
			else:
				exceed_list.append((tsm_bal.getContainer().location, tsm_bal.maxDate()))

	# Report each closure exceedance at the time of the relevant minimum or
	# maximum residual and display a warning to the user.
	if len(exceed_list) > 0:
		exTime = HecTime()
		for item in exceed_list:
			exTime.set(item[1])
			msg += item[0] + ' @ ' + exTime.dateAndTime(4) + '\n'
		msg += "There may be errors in the operations spreadsheet at these locations."
		print msg
		MessageBox.showWarning(msg, "Volume Closure Warning")

	########################
	# Zero-Flow Time Series
	########################

	# Create a daily zero-flow series for boundary locations that require a
	# valid daily DSS flow record even when no flow is prescribed.
	tsmath_zero_flow_day = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1DAY", "0M", 0.0)
	tsmath_zero_flow_day.setUnits("CFS")
	tsmath_zero_flow_day.setType("PER-AVER")
	tsmath_zero_flow_day.setTimeInterval("1DAY")
	tsmath_zero_flow_day.setLocation("ZERO-BY-DAY")
	tsmath_zero_flow_day.setParameterPart("FLOW-ZERO")
	tsmath_zero_flow_day.setVersion(BC_F_part)
	tsm_list.append(tsmath_zero_flow_day)

	# Create the corresponding hourly zero-flow record.
	tsmath_zero_flow_hour = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1HOUR", "0M", 0.0)
	tsmath_zero_flow_hour.setUnits("CFS")
	tsmath_zero_flow_hour.setType("PER-AVER")
	tsmath_zero_flow_hour.setTimeInterval("1Hour")
	tsmath_zero_flow_hour.setLocation("ZERO-BY-HOUR")
	tsmath_zero_flow_hour.setParameterPart("FLOW-ZERO")
	tsmath_zero_flow_hour.setVersion(BC_F_part)
	tsm_list.append(tsmath_zero_flow_hour)

	# Create an hourly zero-gate-count series.
	tsmath_zero_gates_hour = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1HOUR", "0M", 0.0)
	tsmath_zero_gates_hour.setUnits("Count")
	tsmath_zero_gates_hour.setType("INST-VAL")
	tsmath_zero_gates_hour.setTimeInterval("1Hour")
	tsmath_zero_gates_hour.setLocation("ZERO-BY-HOUR")
	tsmath_zero_gates_hour.setParameterPart("GATES-ZERO")
	tsmath_zero_gates_hour.setVersion(BC_F_part)
	tsm_list.append(tsmath_zero_gates_hour)

	# Create a constant hourly series representing one operating gate.
	tsmath_one_gate_hour = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1HOUR", "0M", 1.0)
	tsmath_one_gate_hour.setUnits("Count")
	tsmath_one_gate_hour.setType("INST-VAL")
	tsmath_one_gate_hour.setTimeInterval("1Hour")
	tsmath_one_gate_hour.setLocation("ONE-BY-HOUR")
	tsmath_one_gate_hour.setParameterPart("GATES-ONE")
	tsmath_one_gate_hour.setVersion(BC_F_part)
	tsm_list.append(tsmath_one_gate_hour)

	# Create a constant hourly series representing three operating gates.
	tsmath_three_gate_hour = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1HOUR", "0M", 3.0)
	tsmath_three_gate_hour.setUnits("Count")
	tsmath_three_gate_hour.setType("INST-VAL")
	tsmath_three_gate_hour.setTimeInterval("1Hour")
	tsmath_three_gate_hour.setLocation("THREE-BY-HOUR")
	tsmath_three_gate_hour.setParameterPart("GATES-3")
	tsmath_three_gate_hour.setVersion(BC_F_part)
	tsm_list.append(tsmath_three_gate_hour)

	# Create a constant hourly series representing two operating gates.
	tsmath_two_gate_hour = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1HOUR", "0M", 2.0)
	tsmath_two_gate_hour.setUnits("Count")
	tsmath_two_gate_hour.setType("INST-VAL")
	tsmath_two_gate_hour.setTimeInterval("1Hour")
	tsmath_two_gate_hour.setLocation("TWO-BY-HOUR")
	tsmath_two_gate_hour.setParameterPart("GATES-2")
	tsmath_two_gate_hour.setVersion(BC_F_part)
	tsm_list.append(tsmath_two_gate_hour)

	# Create a constant hourly series representing five operating gates.
	tsmath_five_gates_hour = tsmath.generateRegularIntervalTimeSeries(
		"%s 0000"%(ops_start_date.date(4)),
		"%s 2400"%(end_time.date(4)),
		"1HOUR", "0M", 5.0)
	tsmath_five_gates_hour.setUnits("Count")
	tsmath_five_gates_hour.setType("INST-VAL")
	tsmath_five_gates_hour.setTimeInterval("1Hour")
	tsmath_five_gates_hour.setLocation("FIVE-BY-HOUR")
	tsmath_five_gates_hour.setParameterPart("GATES-FIVE")
	tsmath_five_gates_hour.setVersion(BC_F_part)
	tsm_list.append(tsmath_five_gates_hour)

	# Write every generated TimeSeriesMath record to the boundary-condition DSS
	# file and construct the corresponding location/path map entry.
	for tsmath_item in tsm_list:
		ts_write = hec.heclib.dss.HecTimeSeries()
		ts_write.setDSSFileName(BC_output_DSS_filename)
		tsc = tsmath_item.getData()
		rv_lines.append("%s,%s,%s,%s"%(
			tsc.location, tsc.parameter,
			Project.getCurrentProject().getRelativePath(BC_output_DSS_filename),
			tsc.fullName))
		if DEBUG: print "\t%s"%rv_lines[-1]
		ts_write.write(tsc)
		ts_write.done()

	return rv_lines