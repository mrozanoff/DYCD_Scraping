# DYCD_Scraping
Is comprised of python scripts designed to scrape reports off of the DYCD Connect system in order to save time on the weekly repetitive manual effort in creating summaries of enrollment, attendance, rate of participation and more. There are two main scripts, the Weekly Dashboard, which runs ROP and Attendance scraping, and Unaccounted Attendance which displays sites' unentered attendance, usually run before the two week DYCD lock is set.

# Weekly Dashboard
Creates a single output Excel file with all reports calculations in one document.
# Unaccounted Attendance
Creates directory structure containing current weeks unaccounted attendance reports and summary, and the same for two weeks prior and analyzes the data change after the DYCD data lock.
