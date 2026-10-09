'''
CVP_ops_tools

stuff to process time series data out of Central Valley Progect operations spreadsheets
'''

# HEC-DSS classes provide time handling, containers, and time-series mathematics.
import hec.heclib.util.HecTime as HecTime
import hec.io.TimeSeriesContainer as tscont
import hec.hecmath.TimeSeriesMath as tsmath
import hec.lang.Const

# Java and Apache POI classes support direct reading of legacy and modern Excel workbooks.
import java.lang
import java.io.File
import java.io.FileInputStream

from org.apache.poi.xssf.usermodel import XSSFWorkbook
from org.apache.poi.hssf.usermodel import HSSFWorkbook
from org.apache.poi.ss import usermodel as SSUsermodel

# Enable diagnostic console output throughout spreadsheet parsing and transformations.
DEBUG = True

# Month lookup tables use a dummy element at index 0 so month numbers map
# directly to list indices.
month_TLA = ["NM", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
days_in_month = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

def get_days_in_month(month_int, year_int):
	"""Return the number of days in a specified month and year.

	Parameters
	----------
	month_int : int
		Calendar month number in the range 1 through 12.
	year_int : int
		Calendar year used to determine whether February is in a leap year.

	Returns
	-------
	int
		Number of days in the requested month.

	Raises
	------
	ValueError
		If ``month_int`` is outside the range 1 through 12.
	"""
	if month_int > 12 or month_int < 1:
		raise ValueError("Month (%d) is not an int between 1 and 12."%(month_int))
	if month_int == 2 and HecTime.isLeap(year_int):
		return 29
	return days_in_month[month_int]

def month_index(s_month):
	"""Convert a three-letter month abbreviation to its numeric index.

	Parameters
	----------
	s_month : str
		Month abbreviation, matched case-insensitively against ``month_TLA``.

	Returns
	-------
	int
		Month number from 1 through 12, or 0 when the abbreviation is not found.
	"""
	if not s_month.upper() in month_TLA: return 0
	rv = 0
	for tla in month_TLA:
		if s_month.upper() == tla:
			break
		rv += 1
	return rv

def next_month(index):
	"""Return the month index immediately following the supplied month.

	Parameters
	----------
	index : int
		Current month index.

	Returns
	-------
	int
		Next month index, wrapping from December to January.
	"""
	if index > 11:
		return 1
	else: return index + 1

def previous_month(index):
	"""Return the month index immediately preceding the supplied month.

	Parameters
	----------
	index : int
		Current month index.

	Returns
	-------
	int
		Previous month index, wrapping from January to December.
	"""
	if index < 2:
		return 12
	else: return index - 1

def is_convertable_to_float(input):
	"""Test whether a value can be converted to a floating-point number.

	Parameters
	----------
	input : object
		Value passed to Python's ``float`` conversion.

	Returns
	-------
	bool
		True when conversion succeeds; otherwise False.
	"""
	try:
		test_val = float(input)
		return True
	except:
		return False


def import_CVP_Ops_csv(ops_fname, forecast_locations, active_locations):
	"""Read a CVP operations CSV file and group rows by forecast location.

	Parameters
	----------
	ops_fname : str
		Path to the operations CSV file.
	forecast_locations : sequence of str
		Location names that identify forecast sections in the file.
	active_locations : sequence of str
		Locations for which profile-date information is retained.

	Returns
	-------
	dict
		Mapping from forecast location names to lists containing calendar,
		profile-date, and time-series data rows represented as strings.

	Notes
	-----
	The parser uses the first 26 comma-separated fields when detecting calendar
	rows, matching the layout assumptions documented in the original code.
	"""
	current_location = None
	start_month = None
	first_date_index = -1
	location_count = 0
	ts_count = 0
	data_lines = []
	rv_dictionary = {}
	calendar = ""

	# Scan the CSV sequentially, retaining the most recently identified
	# calendar row for each location block.
	with open(ops_fname) as infile:
		num_lines = 0; num_data_lines = 0
		for line in infile:
			num_lines += 1
			line_contains_months = False
			token = line.strip().split(',')

			# figure out what columns our data start in, what month we're looking at, and ignore blank lines
			# the sample spreadsheet had an unused summary block starting in column AA, which I'm ignoring
			num_t = 0; num_val = 0
			for t in token[:26]:
				if len(t.strip()) > 0:
					num_val += 1
					if not line_contains_months and t.strip().upper() in month_TLA:
						line_contains_months = True
						first_date_index = num_t
						start_month = t.strip().upper()
						if DEBUG: print "Calendar line %s: "%(line)
						if DEBUG: print "Found \"%s\" in column %d"%(t.strip(), num_t + 1)
						calendar = line
				num_t += 1
			if num_val == 0:
				continue # don't include this line in the result

			# A recognized location following a calendar row starts a new
			# forecast-location data block.
			if token[0].strip() in forecast_locations and len(calendar) > 0:
				if location_count > 0:
					rv_dictionary[current_location] = data_lines
				data_lines = []
				current_location = token[0].strip()
				print "setting current location to %s"%(current_location)
				data_lines.append("%d,%s"%(first_date_index, calendar.strip()))

				if current_location in active_locations and len(token[1].strip()) > 1:
					print("PROFILEDATE: %s"%(token[1]))
					data_lines.append("PROFILEDATE: %s"%(token[1]))
					if DEBUG: print "setting profile date to %s at %s"%(token[1], current_location)
				location_count += 1
				calendar = ""
				continue

			# Non-calendar rows within the current block are retained as
			# candidate time-series data.
			if not line_contains_months:
				data_lines.append(line.strip())
				ts_count += 1

	rv_dictionary[current_location] = data_lines #
	print "Found %d forecast locations and %d time series in ops file \n\t%s."%(
		location_count, ts_count, ops_fname)
	return rv_dictionary


def monthFromDateStr(str):
	"""Extract a three-letter month abbreviation from a date string.

	Parameters
	----------
	str : str
		Date-like string containing a whitespace-delimited month abbreviation.

	Returns
	-------
	str or None
		Uppercase month abbreviation when found; otherwise None.
	"""
	month_TLA = ["NM", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

	# Search whitespace-delimited date fields for one of the recognized month
	# abbreviations.
	for token in str.split():
		if token.strip().upper() in month_TLA:
			return token.strip().upper()
	return None

def excel_date_str_2_dmy(date_str):
	"""Convert supported Excel date strings to day-month-year text.

	Parameters
	----------
	date_str : str
		Date represented either as Java-style date text or slash-separated text.

	Returns
	-------
	str or None
		Date formatted as ``day-MON-year`` for a recognized input format.

	Notes
	-----
	For example, ``Wed Jan 01 00:00:00 PST 2025`` or ``1/1/2025`` is
	converted to ``1-JAN-2025``.
	"""
	if DEBUG: print "Excel date: " + date_str

	# Support both Java Date string output and slash-separated spreadsheet date
	# text.
	parts = date_str.split()
	if len(parts) == 6:
		return parts[2] + '-' + parts[1].upper() + '-' + parts[5]
	parts = date_str.split('/')
	if len(parts) == 3:
		return parts[1] + '-' + month_TLA[int(parts[0])] + '-' + parts[2]


def import_CVP_Ops_xls(ops_fname, forecast_locations, active_locations, sheet_number=0):
	"""Read a CVP operations XLS/XLSX worksheet and group rows by location.

	Parameters
	----------
	ops_fname : str
		Path to the Excel operations workbook.
	forecast_locations : sequence of str
		Location names that identify forecast sections in the worksheet.
	active_locations : sequence of str
		Locations for which profile-date information is retained.
	sheet_number : int, optional
		Zero-based workbook sheet index to read.

	Returns
	-------
	dict
		Mapping from forecast location names to calendar/profile metadata and
		time-series rows serialized in comma-separated form.

	Notes
	-----
	Excel formats are decoded by the Apache POI library. Formula cells use
	their cached results because formulas are not evaluated by this routine.

	The implementation expects the Apache POI 3.8 API. Newer versions may
	have API changes. In particular, ``SSUsermodel.Cell.CELL_TYPE_XXX`` is a
	constant in POI 3.8 and part of an enumeration in POI 4.x.

	The original developer referenced the following resources for interpreting
	formula cells and Excel dates:

	https://www.baeldung.com/java-apache-poi-cell-string-value
	https://www.baeldung.com/java-read-dates-excel
	"""
	current_location = None
	start_month = None
	first_date_index = -1
	location_count = 0
	ts_count = 0
	data_lines = []
	rv_dictionary = {}
	calendar = ""

	# Select the Apache POI workbook implementation according to the input file
	# extension while preserving the original exception behavior.
	try:
		if ops_fname.endswith(".xlsx"):
			workbook = XSSFWorkbook(
				java.io.FileInputStream(java.io.File(ops_fname)))
		if ops_fname.endswith(".xls"):
			workbook = HSSFWorkbook(
				java.io.FileInputStream(java.io.File(ops_fname)))
	except Exception as e:
		raise e

	# Open the requested worksheet and prepare POI's formatter for converting
	# cells to the CSV-like representation expected by downstream processing.
	sheet = workbook.getSheetAt(sheet_number)
	formatter = SSUsermodel.DataFormatter(True)
	num_lines = 0; num_data_lines = 0
	for row in sheet.iterator():
		num_lines += 1
		num_cols = 0
		line_contains_months = False
		token = []

		# Decode each populated cell while preserving the worksheet column order.
		for cell in row.cellIterator():
			# This business -- Cell.CELL_TYPE_XXX -- has been revised a couple of times
			# between POI version 3.8 and 4.x. Watch out it doesn't bite us
			cellType = cell.getCellType()

			# Formula cells are represented by their cached result because this
			# routine does not evaluate workbook formulas.
			if cellType == SSUsermodel.Cell.CELL_TYPE_FORMULA:
				cachedType = cell.getCachedFormulaResultType()
				print str(cachedType) + " : " + formatter.formatCellValue(cell)
				if cachedType == SSUsermodel.Cell.CELL_TYPE_NUMERIC:
					if SSUsermodel.DateUtil.isCellDateFormatted(cell):
						# calendar rows have month labels starting in column 2
						if num_cols > 1:
							token.append(monthFromDateStr(str(cell.getDateCellValue())))
						else:
							token.append(str(cell.getDateCellValue()))
					else:
						token.append(str(cell.getNumericCellValue()))
				if cachedType == SSUsermodel.Cell.CELL_TYPE_STRING:
					token.append(str(cell.getStringCellValue()))
			else:
				if cellType == SSUsermodel.Cell.CELL_TYPE_STRING:
					token.append(str(cell.getStringCellValue()))
				elif cellType == SSUsermodel.Cell.CELL_TYPE_NUMERIC:
					if SSUsermodel.DateUtil.isCellDateFormatted(cell):
						# calendar rows have month labels starting in column 2
						if num_cols > 1:
							token.append(monthFromDateStr(str(cell.getDateCellValue())))
						else:
							token.append(str(cell.getDateCellValue()))
					else:
						token.append(str(cell.getNumericCellValue()))
				else:
					token.append(formatter.formatCellValue(cell))
			num_cols += 1

		# figure out what columns our data start in, what month we're looking at, and ignore blank lines
		num_t = 0; num_val = 0

		# Detect calendar rows and record the first month column used by the
		# operations data that follow.
		for t in token:
			if len(t.strip()) > 0:
				num_val += 1
				# if there's a month label in the first 6 cells of the row, the row is a calendar line
				if ((not line_contains_months) and
					num_t < 6 and
					t.strip().upper() in month_TLA):
					line_contains_months = True
					first_date_index = num_t
					start_month = t.strip().upper()
					if DEBUG: print "Calendar line %d: "%(num_lines)
					if DEBUG: print "Found \"%s\" in column %d"%(t.strip(), num_t + 1)
					calendar = ','.join(token)
			num_t += 1

		if num_val == 0:
			continue # don't include this row in the result

		# A recognized location following a calendar row begins a new
		# forecast-location block; save the previous block before replacing it.
		if token[0].strip() in forecast_locations and len(calendar) > 0:
			if location_count > 0:
				rv_dictionary[current_location] = data_lines
				data_lines = []
			current_location = token[0].strip()
			if DEBUG: print "setting current location to %s"%(current_location)
			print "%d,%s"%(first_date_index, calendar)
			data_lines.append("%d,%s"%(first_date_index, calendar))
			if current_location in active_locations and len(token[1].strip()) > 1:
				data_lines.append("PROFILEDATE: %s"%(excel_date_str_2_dmy(token[1])))
				if DEBUG: print "setting profile date to %s at %s"%(excel_date_str_2_dmy(token[1]), current_location)
			location_count += 1
			calendar = ""
			continue

		# Retain sufficiently populated non-calendar rows as time-series data.
		if not line_contains_months and num_val > 10:
			data_lines.append(','.join(token))
			ts_count += 1

	rv_dictionary[current_location] = data_lines #
	print "Found %d forecast locations and %d time series in ops file \n\t%s."%(
		location_count, ts_count, ops_fname)
	return rv_dictionary

def make_ops_tsc(location_name, start_year, start_month, ts_line, data_type=None, data_units=None, ops_label=None, currentAlternative=None):
	"""Convert an operations-spreadsheet row to a monthly time-series container.

	Parameters
	----------
	location_name : str
		Location assigned to the output DSS time series.
	start_year : int
		Year associated with the first monthly value.
	start_month : str
		Three-letter month abbreviation for the first value.
	ts_line : str
		Comma-separated operations row containing a parameter label and values.
	data_type : str, optional
		DSS data type; defaults to ``PER-CUM`` when not supplied.
	data_units : str, optional
		DSS units; defaults to ``AC-FT`` when not supplied.
	ops_label : str, optional
		Version/F-part label assigned to the output DSS pathname.
	currentAlternative : object, optional
		Calling application object used to receive compute messages.

	Returns
	-------
	hec.io.TimeSeriesContainer
		Monthly time-series container populated from the spreadsheet row.

	Notes
	-----
	The input row is assumed to represent a one-month time step. Volumes are
	expected in TAF and are converted to acre-feet, while flows are expected
	in CFS. Parameter text can override the default units and DSS data type.
	"""
	i_year = start_year

	# Report the source row and starting date through the application message
	# interface when available, and through the console when debugging.
	if currentAlternative:
		currentAlternative.addComputeMessage("making a time series at %s starting at %s %d..."%(location_name, start_month, i_year))
		currentAlternative.addComputeMessage("From data line: \"%s\""%(ts_line))
	if DEBUG:
		print "making a time series at %s starting at %s %d..."%(location_name, start_month, i_year)
		print "From data line: \"%s\""%(ts_line)
	param = ""

	# Apply the default DSS interpretation unless the caller explicitly
	# supplies type, units, or version metadata.
	if not data_type:
		data_type = "PER-CUM" # "PER-AVER"
	if not data_units:
		data_units = "AC-FT" # "CFS"
	if not ops_label:
		ops_label=""

	# Initialize the calendar position and value/time arrays used to construct
	# the monthly TimeSeriesContainer.
	i_month = month_index(start_month)
	ts_vals = []
	ts_times = []    #HecTime objects
	ts_minutes = []  #minutes
	t_count = 0
	v_count = 0

	# Parse the row from left to right. Text identifies the parameter, while the
	# contiguous numeric block supplies successive monthly values.
	tokens = ts_line.split(',')
	for token in tokens:
		t_count += 1
		if len(token.strip()) == 0:
			# values are in a continuous block of comma-separated values, so a null
			# after we've started adding values means we've reached the last value
			if v_count > 0:
				break
			# a null before we've started adding values means we're still looking for
			# the first value
			continue

		# first non-empty field is the parameter
		try:
			# Numeric fields become successive monthly values. Values using the
			# default acre-foot interpretation are converted from TAF to acre-feet.
			ts_vals.append(float(token.strip()))
			v_count += 1
			if data_units == "AC-FT":
				ts_vals[-1] = ts_vals[-1]*1000. # convert TAF to Acre-Feet
			dateTime = HecTime()
			dateTime.setYearMonthDay(i_year, i_month, get_days_in_month(i_month, i_year), 1440)
			ts_times.append(dateTime)
			last_month = i_month
			i_month = next_month(i_month)
			if (i_month - last_month) < 0:
				i_year += 1
		except ValueError:
			# conversion to float failed, so the token is assumed to be the paramter (C Part)
			# of the time series unless we've already got one
			if len(param) == 0:
				param = token.strip().upper()
				if currentAlternative:
					currentAlternative.addComputeMessage("making a time series of %s at %s ..."%(param, location_name))

				# Infer DSS units/type from recognized parameter suffixes in the
				# operations spreadsheet label.
				if param.strip(')').endswith("CFS") or param.endswith("AFRP"):
					param = "FLOW-" + param
					data_type = "PER-AVER"
					data_units = "CFS"
				if param.strip(')').endswith("ACRES"):
					data_type = "INST-VAL"
					data_units = "ACRES"
				if param.strip(')').endswith("FEET"):
					data_type = "INST-VAL"
					data_units = "FEET"
				if param.strip(')').endswith("STORAGE"):
					data_type = "INST-VAL"


	# working_math = tsmath.generateRegularIntervalTimeSeries(time_start.date(8), time_end.date(8), "1MON", 1.0)
	# convert HecTimes to minutes
	for dt in ts_times:
		ts_minutes.append(dt.getMinutes())

	# Populate the TimeSeriesContainer values, times, DSS metadata, and monthly
	# pathname from the parsed operations row.
	rv_tsc = tscont()
	rv_tsc.type = data_type
	rv_tsc.units = data_units
	rv_tsc.numberValues = len(ts_vals)
	rv_tsc.values = ts_vals
	rv_tsc.times = ts_minutes
	rv_tsc.startTime = ts_minutes[0]
	rv_tsc.location = location_name.replace('/', '-')
	rv_tsc.parameter = param.replace('/', '-')
	rv_tsc.interval = 43200
	rv_tsc.version = ops_label
	rv_tsc.fullName = "//%s/%s//1MON/%s/"%(rv_tsc.location,rv_tsc.parameter,rv_tsc.version)

	return rv_tsc


def uniform_transform_monthly_to_daily(tsmath_months, start_day_count=None, currentAlternative=None):
	"""Disaggregate monthly values to a uniform daily time series.

	Parameters
	----------
	tsmath_months : hec.hecmath.TimeSeriesMath
		Monthly input series containing volumes, flows, or temperatures.
	start_day_count : int, optional
		Number of days represented by a partial first month.
	currentAlternative : object, optional
		Calling application object used to receive compute messages.

	Returns
	-------
	hec.hecmath.TimeSeriesMath
		Daily series containing uniformly distributed monthly values.

	Notes
	-----
	The input time series is assumed to contain either monthly average daily
	flows or monthly volumes in acre-feet.

	A ``start_day_count`` of *n* means that the output begins with the final
	*n* days represented by the first month in the input. For example, an
	April input with ``start_day_count=14`` produces a daily series beginning
	at the end of April 16 according to the original implementation.

	TAF inputs are converted to acre-feet and then to average CFS. CFS and
	temperature inputs retain their applicable units/type while being repeated
	at the daily interval.
	"""
	# Establish the monthly input extent and derive the number of days
	# represented by the first output month.
	start_time_in = HecTime(tsmath_months.firstValidDate(), HecTime.MINUTE_INCREMENT)
	end_time_in = HecTime(tsmath_months.lastValidDate(), HecTime.MINUTE_INCREMENT)

	if not start_day_count:
		start_day_count = start_time_in.day()

	#output time series will begin start_day_count days before the end of the first month in the monthly input ts
	start_day_of_month = 1 + get_days_in_month(start_time_in.month(), start_time_in.year()) - start_day_count
	start_time_out = HecTime()
	start_time_out.setYearMonthDay(start_time_in.year(), start_time_in.month(), start_day_of_month, 0)
	print "uniform daily time series start time = " + start_time_out.date(4) + ' ' + str(start_time_out.minutesSinceMidnight())

	# Determine whether the monthly input represents volume requiring
	# acre-foot-to-CFS conversion or values that can retain their units.
	input_is_acrefeet = True
	if tsmath_months.getUnits().upper().startswith("TAF"):
		tsmath_months = tsmath_months.multiply(1000.0)
		tsmath_months.setUnits("AC-FT")
		tsmath_months.setType("PER-CUM")
	elif (tsmath_months.getUnits().upper().startswith("CFS") or
		tsmath_months.getUnits().upper().startswith("DEG") or
		(len(tsmath_months.getUnits().upper().strip()) == 1 and
		(tsmath_months.getUnits().upper() == "F" or 
		tsmath_months.getUnits().upper() == "C"))):
		input_is_acrefeet = False

	# get the date and time value lists from the TimeSeriesMath objects
	tsc_months = tsmath_months.getData()

	# Report the transformation through the calling application when available,
	# otherwise use diagnostic console output when DEBUG is enabled.
	if currentAlternative:
		currentAlternative.addComputeMessage("Calculating uniform time series for %s at %s"%(tsc_months.parameter, tsc_months.location))
		currentAlternative.addComputeMessage("Input time series starting at %s"%(str(start_time_in)))
		currentAlternative.addComputeMessage("Output time series starting at %s"%(str(start_time_out)))

	elif DEBUG:
		print "Calculating uniform time series for %s at %s"%(tsc_months.parameter, tsc_months.location)
		print "Input time series starting at %s"%(str(start_time_in))
		print "Output time series starting at %s"%(str(start_time_out))

	# create HecTime objects for indexing the pattern and output time series
	search_time = HecTime()
	post_time = HecTime()

	# Create the daily output grid and update the pathname interval to identify
	# the generated daily record.
	tsc_result = tsmath.generateRegularIntervalTimeSeries(start_time_out.date(8), end_time_in.date(8), "1DAY", "0M", 1.0).getData()
	path_parts = tsc_months.fullName.split('/')
	path_parts[5] = "1DAY"
	tsc_result.fullName = '/'.join(path_parts)
	tsc_result.version = "UNIFORM"
	tsc_result.location = tsc_months.location

	# Acre-foot volumes become period-average CFS; non-volume inputs retain
	# their original units, data type, and parameter.
	if input_is_acrefeet:
		tsc_result.units = "CFS"
		tsc_result.type = "PER-AVER"
	else:
		tsc_result.units = tsc_months.units
		tsc_result.type = tsc_months.type
		tsc_result.parameter = tsc_months.parameter

	i = 0

	# Fill each daily interval from the applicable monthly value, applying the
	# acre-foot-to-average-CFS conversion when required.
	for tm in tsc_result.times:
		post_time.setMinutes(tm)
		# print "post_time = " + post_time.date(4) + ' ' + str(post_time.minutesSinceMidnight())

		cfs_conversion = 1.
		if input_is_acrefeet:
			if tm <= tsc_months.times[0]:
				cfs_conversion = 0.50417/start_day_count
			else:
				cfs_conversion = 0.50417/get_days_in_month(post_time.month(), post_time.year())

		if tm <= tsc_months.times[0]:
			tsc_result.values[i] = tsc_months.values[0]*cfs_conversion
		else:
			search_time.setYearMonthDay(post_time.year(), post_time.month(),
				get_days_in_month(post_time.month(), post_time.year()), 1440)
			tsc_result.values[i] = tsc_months.getValue(search_time)*cfs_conversion

		i += 1

	return tsmath(tsc_result)


def uniform_transform_monthly_to_hourly(tsmath_months, start_day_count=None, currentAlternative=None):
	"""Disaggregate monthly values to a uniform hourly time series.

	Parameters
	----------
	tsmath_months : hec.hecmath.TimeSeriesMath
		Monthly input series containing volumes or flows.
	start_day_count : int, optional
		Number of days represented by a partial first month.
	currentAlternative : object, optional
		Calling application object used to receive compute messages.

	Returns
	-------
	hec.hecmath.TimeSeriesMath
		Hourly series containing uniformly distributed monthly values.

	Notes
	-----
	The input time series is assumed to contain either monthly average daily
	flows or monthly volumes in acre-feet.

	A ``start_day_count`` of *n* indicates that the output begins with the last
	*n* days represented by the first input month. The original developer
	documentation gives an April input with ``start_day_count=14`` as producing
	an output beginning on April 17.

	Monthly acre-foot volumes are converted to average CFS; flow inputs retain
	their existing units and DSS data type.
	"""
	# Establish the monthly input extent and determine the beginning of the
	# potentially partial first output month.
	start_time_in = HecTime(tsmath_months.firstValidDate(), HecTime.MINUTE_INCREMENT)
	end_time_in = HecTime(tsmath_months.lastValidDate(), HecTime.MINUTE_INCREMENT)

	if not start_day_count:
		start_day_count = start_time_in.day()

	#output time series will begin start_day_count days before the end of the first month in the monthly input ts
	start_day_of_month = 1 + get_days_in_month(start_time_in.month(), start_time_in.year()) - start_day_count
	start_time_out = HecTime()
	start_time_out.setYearMonthDay(start_time_in.year(), start_time_in.month(), start_day_of_month, 0)

	print "hourly start time = " + start_time_out.date(4) + ' ' + str(start_time_out.minutesSinceMidnight())

	# Determine whether the input represents monthly volume requiring conversion
	# to average flow or an existing CFS flow series.
	input_is_acrefeet = True
	if tsmath_months.getUnits().upper().startswith("TAF"):
		tsmath_months = tsmath_months.multiply(1000.0)
		tsmath_months.setUnits("AC-FT")
		tsmath_months.setType("PER-CUM")
	elif tsmath_months.getUnits().upper().startswith("CFS"):
		input_is_acrefeet = False

	# get the date and time value lists from the TimeSeriesMath objects
	tsc_months = tsmath_months.getData()

	# Report the transformation through the calling application or DEBUG output.
	if currentAlternative:
		currentAlternative.addComputeMessage("Calculating uniform time series for %s at %s"%(tsc_months.parameter, tsc_months.location))
		currentAlternative.addComputeMessage("Input time series starting at %s"%(str(start_time_in)))
		currentAlternative.addComputeMessage("Output time series starting at %s"%(str(start_time_out)))

	elif DEBUG:
		print "Calculating uniform time series for %s at %s"%(tsc_months.parameter, tsc_months.location)
		print "Input time series starting at %s"%(str(start_time_in))
		print "Output time series starting at %s"%(str(start_time_out))

	# create HecTime objects for indexing the pattern and output time series
	search_time = HecTime()
	post_time = HecTime()

	# Create the hourly output grid and modify the DSS pathname interval to
	# identify the generated hourly record.
	tsc_result = tsmath.generateRegularIntervalTimeSeries(start_time_out.date(8), end_time_in.date(8), "1HOUR", "0M", 1.0).getData()
	path_parts = tsc_months.fullName.split('/')
	path_parts[5] = "1HOUR"
	tsc_result.fullName = '/'.join(path_parts)
	tsc_result.version = "UNIFORM"
	tsc_result.location = tsc_months.location

	# Volume inputs become period-average CFS; existing flow inputs preserve
	# their original units, type, and parameter metadata.
	if input_is_acrefeet:
		tsc_result.units = "CFS"
		tsc_result.type = "PER-AVER"
	else:
		tsc_result.units = tsc_months.units
		tsc_result.type = tsc_months.type
		tsc_result.parameter = tsc_months.parameter

	i = 0

	# Populate each hourly interval from the corresponding monthly quantity,
	# applying the same monthly volume-to-CFS factor used by the daily routine.
	for tm in tsc_result.times:
		post_time.setMinutes(tm)
		# print "post_time = " + post_time.date(4) + ' ' + str(post_time.minutesSinceMidnight())

		cfs_conversion = 1.
		if input_is_acrefeet:
			if tm <= tsc_months.times[0]:
				cfs_conversion = 0.50417/start_day_count
			else:
				cfs_conversion = 0.50417/get_days_in_month(post_time.month(), post_time.year())

		if tm <= tsc_months.times[0]:
			tsc_result.values[i] = tsc_months.values[0]*cfs_conversion
		else:
			search_time.setYearMonthDay(post_time.year(), post_time.month(),
				get_days_in_month(post_time.month(), post_time.year()), 1440)
			tsc_result.values[i] = tsc_months.getValue(search_time)*cfs_conversion

		i += 1

	return tsmath(tsc_result)

def weight_transform_monthly_to_daily(tsmath_months, tsmath_pattern, start_day_count=None, currentAlternative=None):
	"""Disaggregate monthly values to daily flows using an annual pattern.

	Parameters
	----------
	tsmath_months : hec.hecmath.TimeSeriesMath
		Monthly input volumes or average flows to be disaggregated.
	tsmath_pattern : hec.hecmath.TimeSeriesMath
		Daily pattern series defining relative within-month flow variation.
	start_day_count : int, optional
		Number of days represented by a partial first month.
	currentAlternative : object, optional
		Calling application object used to receive compute messages.

	Returns
	-------
	hec.hecmath.TimeSeriesMath
		Daily average-flow series scaled to the monthly input values.

	Notes
	-----
	The pattern time series is assumed to cover a calendar year and to use a
	non-leap year; the original developer notes identify year 3000 as suitable.
	Both the pattern and output time series are daily average flows in CFS.

	The monthly input is assumed to contain either monthly average daily flows
	or monthly volumes in acre-feet. Monthly scale factors preserve the monthly
	magnitude while retaining the daily shape of the supplied pattern. A
	partial first month receives a separate scale calculation using only the
	represented pattern days.
	"""
	# Establish the input and pattern time extents used to align monthly values
	# with the recurring daily distribution pattern.
	start_time_in = HecTime(tsmath_months.firstValidDate(), HecTime.MINUTE_INCREMENT)
	end_time_in = HecTime(tsmath_months.lastValidDate(), HecTime.MINUTE_INCREMENT)
	start_time_pattern = HecTime(tsmath_pattern.firstValidDate(), HecTime.MINUTE_INCREMENT)

	if not start_day_count:
		start_day_count = start_time_in.day()

	#output time series will begin start_day_count days before the end of the first month in the monthly input ts
	start_day_of_month = 1 + get_days_in_month(start_time_in.month(), start_time_in.year()) - start_day_count
	start_time_out = HecTime()
	start_time_out.setYearMonthDay(start_time_in.year(), start_time_in.month(), start_day_of_month, 0)

	# Determine whether the monthly input represents volume requiring conversion
	# to flow or already contains average flow values.
	input_is_acrefeet = True
	if tsmath_months.getUnits().upper().startswith("TAF"):
		tsmath_months = tsmath_months.multiply(1000.0)
		tsmath_months.setUnits("AC-FT")
		tsmath_months.setType("PER-CUM")
	elif tsmath_months.getUnits().upper().startswith("CFS"):
		input_is_acrefeet = False

	# calculate monthly average daily flows from the daily flow pattern time series
	tsmath_pattern_ave = tsmath_pattern.transformTimeSeries("1MON", "", "AVE")

	# get the date and time value lists from the TimeSeriesMath objects
	tsc_months = tsmath_months.getData()
	tsc_pattern = tsmath_pattern.getData()
	tsc_pattern_ave = tsmath_pattern_ave.getData()

	# Report the weighted transformation through the calling application when
	# available, or through DEBUG console output otherwise.
	if currentAlternative:
		currentAlternative.addComputeMessage("Calculating weighted time series for %s at %s"%(tsc_months.parameter, tsc_months.location))
		currentAlternative.addComputeMessage("Input time series starting at %s"%(str(start_time_in)))
		currentAlternative.addComputeMessage("Output time series starting at %s"%(str(start_time_out)))
	elif DEBUG:
		print "Calculating weighted time series for %s at %s"%(tsc_months.parameter, tsc_months.location)
		print "Input time series starting at %s"%(str(start_time_in))
		print "Output time series starting at %s"%(str(start_time_out))

	# create HecTime objects for indexing the pattern and output time series
	search_time = HecTime()
	post_time = HecTime()

	# Make a dictionary of volume ratios by month (i.e. this month's volume/pattern year volume for month)
	scale_lookup = {}
	in_time = HecTime( HecTime.MINUTE_INCREMENT)

	# Calculate a scale factor for every input year-month by comparing the
	# monthly operations value with the corresponding pattern-month magnitude.
	for time_int in tsc_months.times:
		in_time.set(time_int)
		print "Input date: %d %s %d (%d)"%(in_time.day(), month_TLA[in_time.month()], in_time.year(), time_int)
		search_time.setYearMonthDay(start_time_pattern.year(), in_time.month(), days_in_month[in_time.month()], 1440)

		key = in_time.year()*100+in_time.month()

		if input_is_acrefeet:
			scale_lookup[key] = tsc_months.getValue(in_time)*0.50417/days_in_month[in_time.month()]/tsc_pattern_ave.getValue(search_time)
		else:
			scale_lookup[key] = tsc_month.getValue(in_time)/tsc_pattern_ave.getValue(search_time)

		if currentAlternative:
			currentAlternative.addComputeMessage("scale for %s %d = %f"%(month_TLA[in_time.month()], in_time.year(), scale_lookup[in_time.month()]))
		elif DEBUG:
			print "scale for %s %d = %f"%(month_TLA[in_time.month()], in_time.year(), scale_lookup[key])

	# if we're starting mid-month, recalculate acre-feet scale factor for the first month
	if input_is_acrefeet and start_day_of_month > 1:
		first_month_key = start_time_in.year()*100 + start_time_in.month()
		sum_flows = 0.0
		i = 0
		search_time.setYearMonthDay(start_time_pattern.year(), start_time_in.month(), start_day_of_month, 1440)
		first_month = search_time.month()

		# Sum only the pattern days represented by the partial first month so
		# its scale factor preserves the supplied partial-month volume.
		if DEBUG: print "Starting pattern time series at %s"%search_time.date(4)
		while search_time.month() == first_month:
			sum_flows += tsc_pattern.getValue(search_time)
			i += 1
			print "index = %d"%(i)
			search_time.addDays(1)
		scale_lookup[first_month_key] = tsc_months.values[0]*0.50417 / sum_flows
		if DEBUG: print "{}AF/{}cfs-day = {}".format(tsc_months.values[0], sum_flows, scale_lookup[first_month_key])

	# if we're starting first-of-month, duplicate acre-feet scale factor for the first month to the previous month
	if input_is_acrefeet and start_day_of_month == 1:
		key = 0
		if start_time_in.month() == 1:
			key = (start_time_in.year() - 1)*100 + 12
		else:
			key = start_time_in.year()*100 + start_time_in.month() - 1
		first_month_key = start_time_in.year()*100 + start_time_in.month()
		scale_lookup[key] = scale_lookup[first_month_key]

	# Create the daily result container while retaining the original monthly
	# pathname; its version is subsequently identified as a weighted result.
	tsc_result = tsmath.generateRegularIntervalTimeSeries(start_time_out.date(8), end_time_in.date(8), "1DAY", "0M", 1.0).getData()
	tsc_result.fullName = tsc_months.fullName
	tsc_result.units = "CFS"
	tsc_result.type = "PER-AVER"

	i = 0

	# Apply the appropriate year-month scale factor to each corresponding day
	# in the recurring reference pattern.
	for time_min in tsc_result.times:
		post_time.setMinutes(time_min)
		search_time.setYearMonthDay(start_time_pattern.year(), post_time.month(), post_time.day(), 1440)
		scale = scale_lookup[post_time.month()+100*post_time.year()]
		tsc_result.values[i] = scale * tsc_pattern.getValue(search_time)
		i += 1

	tsm_result = tsmath(tsc_result)
	tsm_result.setVersion("WEIGHTED")
	print "Weight disaggregation of %s complete."%(tsc_result.fullName)
	return tsm_result


def split_time_series_static(tsmath_in, names_weights, out_param_name):
	"""Split a time series among locations using constant relative weights.

	Parameters
	----------
	tsmath_in : hec.hecmath.TimeSeriesMath
		Input time series to partition.
	names_weights : dict
		Mapping of output location names to static numeric weights.
	out_param_name : str
		Parameter name assigned to each output series.

	Returns
	-------
	list
		TimeSeriesMath objects containing the normalized weighted partitions.

	Notes
	-----
	The supplied weights are normalized at computation time so the resulting
	component time series sum to the original input series.
	"""
	rv_tsmath_list = []

	# Sum all supplied weights before calculating each location's normalized
	# fraction of the original time series.
	total_weight = 0.
	for key in names_weights.keys():
		total_weight += names_weights[key]

	# Generate one output series for each configured location.
	for key in names_weights:
		tsmath_product = tsmath_in.multiply(names_weights[key]/total_weight)
		tsmath_product.setParameterPart(out_param_name)
		tsmath_product.setLocation(key)
		rv_tsmath_list.append(tsmath_product)

	return rv_tsmath_list


def split_time_series_monthly(tsmath_in, names_weights, out_param_name):
	"""Split a time series using location weights that vary by month.

	Parameters
	----------
	tsmath_in : hec.hecmath.TimeSeriesMath
		Input time series to partition.
	names_weights : dict
		Mapping of location names to 12 monthly weights ordered January through
		December.
	out_param_name : str
		Parameter name assigned to each output series.

	Returns
	-------
	list
		TimeSeriesMath objects containing normalized month-dependent partitions.

	Notes
	-----
	The monthly weights are normalized at computation time. Each dictionary
	value is expected to contain 12 weights, one for each month from January
	through December.
	"""
	rv_tsmath_list = []

	# Calculate a separate normalization denominator for each calendar month.
	total_weight = []
	for i in range(12):
		month_sum = 0
		for key in names_weights.keys():
			month_sum += names_weights[key][i]
		total_weight.append(month_sum)

	# Create a daily weighting container spanning the same valid period as the
	# input time series.
	time_start = HecTime(tsmath_in.firstValidDate(), HecTime.MINUTE_INCREMENT)
	time_end = HecTime(tsmath_in.lastValidDate(), HecTime.MINUTE_INCREMENT)
	weight_container = tsmath.generateRegularIntervalTimeSeries(
		time_start.date(8), time_end.date(8), "1DAY", "0M", 1.0).getData()

	# For each output location, populate the daily weighting series with the
	# normalized weight corresponding to each value's calendar month.
	for key in names_weights.keys():
		for i in range(weight_container.numberValues):
			time_end.set(weight_container.times[i])
			weight_container.values[i] = names_weights[key][time_end.month() - 1]/total_weight[time_end.month() - 1]
		weight_math = tsmath(weight_container)
		tsmath_product = tsmath_in.multiply(weight_math)
		tsmath_product.setParameterPart(out_param_name)
		tsmath_product.setLocation(key)
		rv_tsmath_list.append(tsmath_product)

	return rv_tsmath_list


def backwardsMovingAverage(tsmath_in, num_periods):
	"""Calculate a backward-looking moving average of a time series.

	Parameters
	----------
	tsmath_in : hec.hecmath.TimeSeriesMath
		Input time series.
	num_periods : int
		Maximum number of current-and-prior values included in each average.

	Returns
	-------
	hec.hecmath.TimeSeriesMath
		Time series containing the backward moving-average values.

	Notes
	-----
	This routine exists because the original developer noted that DSSMath did
	not provide the required backward moving-average function. HEC undefined
	double values are excluded from both the moving sum and value count.
	"""
	# Obtain a copy for the result while retaining direct access to the original
	# input values for moving-window calculations.
	rv_tsc = tsmath_in.getData() # getData() returns a copy of the tsMath's time-series container
	rv_parts = rv_tsc.fullName.strip('/').split('/')
	i = 0; j = 0
	in_vals = tsmath_in.getContainer().values # getContainer() returns access to the time-series container in place

	if DEBUG:
		print "Input TSMath for moving average contains %d values."%(tsmath_in.getContainer().numberValues)

	out_vals =[]

	# For each input position, average the current value and available preceding
	# values within the requested window while ignoring DSS undefined values.
	for val in in_vals:
		j += 1
		k = j - num_periods
		moving_sum = 0.
		moving_count = 0.
		if k < 0: k = 0
		for addend in in_vals[k:j]:
			if addend == hec.lang.Const.UNDEFINED_DOUBLE:
				continue
			else:
				moving_sum += addend
				moving_count += 1.0
		out_vals.append(moving_sum/moving_count)
		i += 1

	# Replace the copied values with the calculated moving averages and assign
	# the pathname used by the original implementation.
	if DEBUG:
		print "Result TSMath for moving average contains %d values."%(len(out_vals))
	rv_tsc.values = out_vals
	rv_tsc.fullName = "//test/flow-avg//" + rv_parts[-2] + "/moving/"
	return tsmath(rv_tsc)

def evaluate_temp_regression(tsmath_flow, tsmath_airtemp, temp_regression_coefficients, currentAlternative = None):
	"""Estimate water temperature from flow and air-temperature regressors.

	Parameters
	----------
	tsmath_flow : hec.hecmath.TimeSeriesMath
		Flow time series used by the regression; calculations require CFS.
	tsmath_airtemp : hec.hecmath.TimeSeriesMath
		Air-temperature series; calculations require degrees Celsius.
	temp_regression_coefficients : sequence of float
		Regression coefficients ordered as intercept, flow coefficient,
		air-temperature coefficient, and regression error statistic.
	currentAlternative : object, optional
		Calling application object used to receive compute messages.

	Returns
	-------
	hec.hecmath.TimeSeriesMath
		Hourly interpolated water-temperature series in degrees Celsius.

	Notes
	-----
	The original developer notes specify that both flow and air temperature
	must first be smoothed with a seven-day centered average. Flow is expected
	in CFS, air temperature in degrees Celsius, and the resulting water
	temperature is in degrees Celsius.

	The regression coefficients are interpreted as an intercept, flow
	coefficient, air-temperature coefficient, and RMS error statistic. The RMS
	error value is supplied with the coefficients but is not directly applied
	by this routine.
	"""
	# Report the location being processed through the calling application when
	# that interface is available.
	if currentAlternative: currentAlternative.addComputeMessage("Calculating water temperatures at %s..."%(tsmath_flow.getContainer().location))

	# Convert metric flow to English units so the flow predictor is expressed
	# in the CFS units expected by the regression coefficients.
	if tsmath_flow.isMetric():
		if currentAlternative: currentAlternative.addComputeMessage("Flow units were \"%s\.\""%(tsmath_flow.getUnits()))
		tsmath_flow = tsmath_flow.convertToEnglishUnits()
		if currentAlternative: currentAlternative.addComputeMessage("Flow units converted to \"%s\.\""%(tsmath_flow.getUnits()))

	# Convert English air-temperature input to metric units before applying
	# coefficients calibrated for degrees Celsius.
	if tsmath_airtemp.isEnglish():
		if currentAlternative: currentAlternative.addComputeMessage("Temperature units were \"%s\.\""%(tsmath_airtemp.getUnits()))
		tsmath_airtemp = tsmath_airtemp.convertToMetricUnits()
		if currentAlternative: currentAlternative.addComputeMessage("Temperature units converted to \"%s\.\""%(tsmath_airtemp.getUnits()))

	if DEBUG: print "Calculating temperatures at %s"%(tsmath_flow.getContainer().location)

	# Transform air temperature to daily averages and normalize its temperature
	# unit metadata to degrees Celsius.
	tsmath_airtemp = tsmath_airtemp.transformTimeSeries("1DAY", "", "AVE")
	if "F" in tsmath_airtemp.getUnits().upper():
		if DEBUG: print "Converting temperatures at %s to Celsius."%(tsmath_flow.getContainer().location)
		tsmath_airtemp.setUnits("deg F")
		tsmath_airtemp = tsmath_airtemp.convertToMetricUnits()
	if "C" in tsmath_airtemp.getUnits().upper():
		tsmath_airtemp.setUnits("deg C")

	# Apply the flow term using the required seven-day centered moving average.
	if DEBUG: print "\tApplying flows..."
	tsmath_watertemp = tsmath_flow.centeredMovingAverage(7, False, True).multiply(temp_regression_coefficients[1])
	tsmath_watertemp.setUnits("deg C")

	# Add the air-temperature term, also based on a seven-day centered moving
	# average, followed by the regression intercept.
	if DEBUG: print "\tApplying air temperature..."
	tsmath_watertemp = tsmath_watertemp.add(tsmath_airtemp.centeredMovingAverage(7, False, True).multiply(temp_regression_coefficients[2]))
	if DEBUG: print "\tApplying constant..."
	tsmath_watertemp = tsmath_watertemp.add(temp_regression_coefficients[0])

	# Establish the output period from the valid daily air-temperature record.
	start_time = HecTime(tsmath_airtemp.firstValidDate(), HecTime.MINUTE_INCREMENT)
	start_time.setTime("0000")
	end_time = HecTime(tsmath_airtemp.lastValidDate(), HecTime.MINUTE_INCREMENT)
	# print "Starts at " + start_time.dateAndTime(4)
	# print "Ends at " + end_time.dateAndTime(4)

	# Interpolate the daily regression result onto an hourly grid and identify
	# the resulting series as water temperature.
	tsmath_out = tsmath.generateRegularIntervalTimeSeries(start_time.dateAndTime(4), end_time.dateAndTime(4), "1HOUR", "", 0.0)
	tsmath_out.setUnits("deg C")
	tsmath_out = tsmath_watertemp.transformTimeSeries(tsmath_out, "INT")
	tsmath_out.setParameterPart("TEMP-WATER")

	return tsmath_out


def leapYearTest(currentAlternative):
	"""Report HEC time values around 29 February for year 3000.

	Parameters
	----------
	currentAlternative : object
		Calling application object that receives diagnostic compute messages.

	Returns
	-------
	None
		This diagnostic routine reports computed minute values and returns
		nothing.

	Notes
	-----
	The routine exercises the HEC time implementation around February 29 of
	year 3000. This is relevant to the annual-pattern assumptions elsewhere
	in this module, where the original developer identified year 3000 as a
	suitable non-leap pattern year.
	"""
	# Exercise HecTime around the February/March boundary and report the
	# resulting internal minute values through the calling application.
	test = HecTime(HecTime.MINUTE_INCREMENT)
	test.setYearMonthDay(3000, 2, 28, 1440)
	currentAlternative.addComputeMessage("HecTime 28 Feb 3000 = %d"%(test.getMinutes()))
	test.setYearMonthDay(3000, 2, 29, 1440)
	currentAlternative.addComputeMessage("HecTime 29 Feb 3000 = %d"%(test.getMinutes()))
	test.setYearMonthDay(3000, 3, 1, 1440)
	currentAlternative.addComputeMessage("HecTime 1 Mar 3000 = %d"%(test.getMinutes()))
	return