============
Excel Report
============

Excel report is a simple Django app to export excel report.


Quick Start
------------

1. Add "reportGenerator" to your INSTALLED_APPS setting like this:

	INSTALLED_APPS = [
		...
		'reportGenerator',
	]

2. Include the reportGenerator URLconf in your project urls.py like this:

	path('excel_export/', include('reportGenerator.urls')),
