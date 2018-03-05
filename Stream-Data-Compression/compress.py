import zipfile
import time
import os
import sys
from datetime import datetime, timedelta

"""
This script compresses all files in srcPath and places the zipped archive (by day) in dstPath.

Assumptions:
	- file formats are:
		- StreamData: SCEP_YYMMDDhhmm.dst
		- 90minute: SCEC_YYMMDDhhmm.dst
		- 120hour: SCEW_YYMMDDhhmm.dst
		- 90hour: SCEM_YYMMDDhhmm.dst
		- 24hour: SCEK_YYMMDDhhmm.dst
		- 9hour: SCED_YYMMDDhhmm.dst
		- MonolithPMU: SCET_YYMMDDhhmm.dst
	- ZIP file formats are:
		- StreamData: AStreamData_MMDD.zip
		- 90minute: A90minData_YYMM.zip
		- 120hour: A1min_120hours_YYYY.zip
		- 90hour: A1min_90hours_YYYY.zip
		- 24hour: A15sec_24hour_YYYY.zip
		- 9hour: A6sec_9hour_YYYY.zip
		- MonolithPMU: Monolith_MMDD.zip
"""

def progressBar(progress):
	sys.stdout.write("\r[" + "*" * round(50 * progress) + " " * (50 - round(50 * progress)) + "]     " + "{}%".format(round(100 * progress)))
	sys.stdout.flush()

def zipFiles(srcPath, dstPath, fileType):
	"""
	INPUT:
		srcPath: source Path
		dstPath: destination Path
		fileType: "StreamData", "90minute", "120hour", "90hour", "24hour", "9hour", "Monolith"
	"""
	zipped = {}
	num_files = {}
	fileTypes = ["StreamData", "90minute", "120hour", "90hour", "24hour", "9hour", "Monolith"]

	if fileType not in fileTypes:
		print("Invalid file type. File can only be of types: " + str(fileTypes))
		sys.exit(0)
	if fileType == "StreamData" or fileType == "90minute":
		for monthlyFolder in os.listdir(srcPath):
			if monthlyFolder[0:2] == "17" and int(monthlyFolder[2:4]) > 0:
				print("\r\n\r\nCreating compressed archives for " + monthlyFolder + "...\r\n")
				month = monthlyFolder[2:4]
				currentFolder = os.path.join(srcPath, monthlyFolder)
				filesToCompress = sorted(os.listdir(currentFolder))

				if fileType == "StreamData":
					for day in range(1, 32):
						date = month + str(day).zfill(2)
						zipped[date] = zipfile.ZipFile(os.path.join(dstPath, "AStreamData_" + date + ".zip"), 'w', zipfile.ZIP_DEFLATED)
						num_files[date] = 0

					current = 0
					total = len(filesToCompress)

					for file in filesToCompress:
						currentDay = file[9:11]
						currentDate = month + currentDay
						zipped[currentDate].write(os.path.join(currentFolder, file), file)
						num_files[currentDate] += 1

						# Console display in cmd
						current += 1
						progress = current / total
						progressBar(progress)

					for day in range(1, 32):
						date = month + str(day).zfill(2)
						zipped[date].close()
						if num_files[date] == 0:
							os.remove(os.path.join(dstPath, "AStreamData_" + date + ".zip"))

				elif fileType == "90minute":
					date = "17" + month
					compressed = zipfile.ZipFile(os.path.join(dstPath, "A90minData_" + date + ".zip"), 'w', zipfile.ZIP_DEFLATED)

					current = 0
					total = len(filesToCompress)

					for file in filesToCompress:
						compressed.write(os.path.join(currentFolder, file), file)

						# Console display in cmd
						current += 1
						progress = current / total
						progressBar(progress)

					compressed.close()

	elif fileType == "Monolith":
		for monthlyFolder in os.listdir(srcPath):
			if monthlyFolder[13:15] == "17" and int(monthlyFolder[15:17]) > 0:
				print("\r\n\r\nCreating compressed archives for " + monthlyFolder + "...\r\n")
				month = monthlyFolder[15:17]
				currentFolder = os.path.join(srcPath, monthlyFolder)
				filesToCompress = sorted(os.listdir(currentFolder))

			for day in range(1, 32):
				date = month + str(day).zfill(2)
				zipped[date] = zipfile.ZipFile(os.path.join(dstPath, "Monolith_" + date + ".zip"), 'w', zipfile.ZIP_DEFLATED)
				num_files[date] = 0

			current = 0
			total = len(filesToCompress)

			for file in filesToCompress:
				currentDay = file[9:11]
				currentDate = month + currentDay
				zipped[currentDate].write(os.path.join(currentFolder, file), file)
				num_files[currentDate] += 1

				# Console display in cmd
				current += 1
				progress = current / total
				progressBar(progress)

			for day in range(1, 32):
				date = month + str(day).zfill(2)
				zipped[date].close()
				if num_files[date] == 0:
					os.remove(os.path.join(dstPath, "Monolith_" + date + ".zip"))

	elif fileType == "120hour" or fileType == "90hour" or fileType == "24hour" or fileType == "9hour":
		filesToCompress = sorted(os.listdir(srcPath))

		for year in range(2, 18):
			num_files[year] = 0

			print("\r\n\r\nCreating compressed archives for 20" + str(year).zfill(2) + "...\r\n")
			if fileType == "120hour":
				zipped[year] = zipfile.ZipFile(os.path.join(dstPath, "A1min_120hours_20" + str(year).zfill(2) + ".zip"), 'w', zipfile.ZIP_DEFLATED)

			elif fileType == "90hour":
				zipped[year] = zipfile.ZipFile(os.path.join(dstPath, "A1min_90hours_20" + str(year).zfill(2) + ".zip"), 'w', zipfile.ZIP_DEFLATED)

			elif fileType == "24hour":
				zipped[year] = zipfile.ZipFile(os.path.join(dstPath, "A15sec_24hour_20" + str(year).zfill(2) + ".zip"), 'w', zipfile.ZIP_DEFLATED)

			elif fileType == "9hour":
				zipped[year] = zipfile.ZipFile(os.path.join(dstPath, "A6sec_9hour_20" + str(year).zfill(2) + ".zip"), 'w', zipfile.ZIP_DEFLATED)

		current = 0
		total = len(filesToCompress)

		for file in filesToCompress:
			current_year = int(file[5:7])
			zipped[current_year].write(os.path.join(srcPath, file), file)
			num_files[current_year] += 1

			# Console display in cmd
			current += 1
			progress = current / total
			progressBar(progress)

		for year in range(2, 18):
			zipped[year].close()
			if num_files[year] == 0:
				if fileType == "120hour":
					os.remove(os.path.join(dstPath, "A1min_120hours_20" + str(year).zfill(2) + ".zip"))

				elif fileType == "90hour":
					os.remove(os.path.join(dstPath, "A1min_90hours_20" + str(year).zfill(2) + ".zip"))

				elif fileType == "24hour":
					os.remove(os.path.join(dstPath, "A15sec_24hour_20" + str(year).zfill(2) + ".zip"))

				elif fileType == "9hour":
					os.remove(os.path.join(dstPath, "A6sec_9hour_20" + str(year).zfill(2) + ".zip"))

	print("\r\n\r\nCompression complete!\r\n")

def unzipFiles(dstPath, checkPath):
	print("Decompressing archives...\r\n")

	current = 0
	total = len(os.listdir(dstPath))
	progress = current / total
	progressBar(progress)

	for archive in os.listdir(dstPath):
		unzip = zipfile.ZipFile(os.path.join(dstPath, archive), 'r')
		unzip.extractall(checkPath)
		unzip.close()

		# Console display in cmd
		current += 1
		progress = current / total
		progressBar(progress)

	print("\r\n\r\nDecompression complete!\r\n")

def checkZipNames(dstPath):
	incorrect = 0
	for zippedFolder in os.listdir(dstPath):
		zippedDate = zippedFolder[12:16]
		fileNames = zipfile.ZipFile(os.path.join(dstPath, zippedFolder), 'r').namelist()
		for file in fileNames:
			fileDate = file[7:11]
			if fileDate != zippedDate:
				incorrect += 1
				print("Mismatched date for " + file + " in " + zippedFolder + "!")
	print(str(incorrect) + " mismatches in file/folder names!")

if __name__ == "__main__":
	srcPath = r'D:\24-hr files'
	dstPath = r'D:\24 Hour Files Compressed'
	
	zipFiles(srcPath, dstPath, "24hour")
	# unzipFiles(dstPath, checkPath)
	# checkZipNames(dstPath)