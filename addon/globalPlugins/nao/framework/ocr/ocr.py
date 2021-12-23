#Nao (NVDA Advanced OCR) is an addon that improves the standard OCR capabilities that NVDA provides on modern Windows versions.
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.
#Last update 2021-12-22
#Copyright (C) 2021 Alessandro Albano, Davide De Carne and Simone Dal Maso

import api
import wx
import time
import queueHandler
from logHandler import log
from contentRecog import recogUi, LinesWordsResult
from .ocr_service import OCRService
from .. speech import speech
from .. generic import screen
from .. import language

language.initTranslation()

class OCRResultPageOffset():
	def __init__(self, start, length):
		self.start = start
		self.end = start + length

class OCR:
	def recognize_screenshot(on_start=None, on_finish=None, on_finish_arg=None):
		if isinstance(api.getFocusObject(), recogUi.RecogResultNVDAObject):
			# Translators: Reported when content recognition (e.g. OCR) is attempted,
			# but the user is already reading a content recognition result.
			speech.message(_N("Already in a content recognition result"))
			return False
		if not OCRService.is_uwp_ocr_available():
			# Translators: Reported when Windows OCR is not available.
			speech.message(_N("Windows OCR not available"))
			return False
		if screen.have_curtain():
			# Translators: Reported when screen curtain is enabled.
			speech.message(_N("Please disable screen curtain before using Windows OCR."))
			return False
		if on_start:
			on_start()
		pixels, width, height = screen.take_snapshot_pixels()
		def h(result):
			if isinstance(result, Exception):
				# Translators: Reporting when recognition (e.g. OCR) fails.
				log.error(_N("Recognition failed") + ': ' + str(result))
				speech.queue_message(_N("Recognition failed"))
				if on_finish:
					if on_finish_arg is None:
						on_finish(success=False)
					else:
						on_finish(success=False, arg=on_finish_arg)
				return
			recogUi._recogOnResult(result)
			if on_finish:
				if on_finish_arg is None:
					on_finish(success=True)
				else:
					on_finish(success=True, arg=on_finish_arg)
		OCRService().push_pixels(pixels, width, height, h)
		return True

	def __init__(self):
		self.clear()

	def clear(self):
		self.service = None
		self.source_file_list = []
		self.source_count = 0
		self.results = []
		self.pages_offset = []
		self.on_finish = None
		self.on_finish_arg = None
		self.on_progress = None
		self.progress_timeout = 1
		self.last_progress = 0
		self.source_file = None
		self.must_abort = False

	def abort(self):
		self.must_abort = True

	def recognize_files(self, source_file, source_file_list, on_start=None, on_finish=None, on_finish_arg=None, on_progress=None, progress_timeout=0):
		if not OCRService.is_uwp_ocr_available():
			# Translators: Reported when Windows OCR is not available.
			speech.queue_message(_N("Windows OCR not available"))
			if on_finish:
				if on_finish_arg is None:
					on_finish(source_file=source_file, result=None, pages_offset=None)
				else:
					on_finish(source_file=source_file, result=None, pages_offset=None, arg=on_finish_arg)
			return
		self.clear()
		self.source_file = source_file
		self.on_finish = on_finish
		self.on_finish_arg = on_finish_arg
		self.on_progress = on_progress
		self.progress_timeout = progress_timeout
		self.source_count = len(source_file_list)
		self.source_file_list = source_file_list
		self.last_progress = time.time()
		if on_start:
			on_start(source_file=self.source_file)
		if not self._recognize_next_page():
			if on_finish:
				if on_finish_arg is None:
					on_finish(source_file=source_file, result=None, pages_offset=None)
				else:
					on_finish(source_file=source_file, result=None, pages_offset=None, arg=on_finish_arg)
			self.clear()

	def _recognize_next_page(self):
		if self.must_abort: return False
		remaining = len(self.source_file_list)
		now = time.time()
		if now - self.last_progress >= self.progress_timeout:
			self.last_progress = now
			if self.on_progress:
				self.on_progress(self.source_count - remaining, self.source_count)
		if remaining > 0:
			if not self.service:
				self.service = OCRService()
			self.service.push_bitmap(wx.Bitmap(self.source_file_list.pop(0)), self._on_recognize_result)
			return True
		return False

	def _on_recognize_result(self, result):
		# This might get called from a background thread, so any UI calls must be queued to the main thread.
		if isinstance(result, Exception):
			# Translators: Reporting when recognition (e.g. OCR) fails.
			log.error(_N("Recognition failed") + ': ' + str(result))
			speech.queue_message(_N("Recognition failed"))
			if self.on_finish:
				if self.on_finish_arg is None:
					self.on_finish(source_file=self.source_file, result=result, pages_offset=None)
				else:
					self.on_finish(source_file=self.source_file, result=result, pages_offset=None, arg=self.on_finish_arg)
			self.clear()
			return
		
		if len(self.pages_offset) == 0:
			self.pages_offset.append(OCRResultPageOffset(0, result.textLen))
		else:
			self.pages_offset.append(OCRResultPageOffset(self.pages_offset[len(self.pages_offset) - 1].end, result.textLen))
		
		# Result is a LinesWordsResult, we store all pages data objects that we will merge later in a single LinesWordsResult
		for line in result.data:
			self.results.append(line)
		
		if not self._recognize_next_page():
			# No more pages
			def h():
				if self.on_finish:
					if self.must_abort:
						if self.on_finish_arg is None:
							self.on_finish(source_file=self.source_file, result=None, pages_offset=None)
						else:
							self.on_finish(source_file=self.source_file, result=None, pages_offset=None, arg=self.on_finish_arg)
					else:
						if self.on_finish_arg is None:
							self.on_finish(source_file=self.source_file, result=LinesWordsResult(self.results, result.imageInfo), pages_offset=self.pages_offset)
						else:
							self.on_finish(source_file=self.source_file, result=LinesWordsResult(self.results, result.imageInfo), pages_offset=self.pages_offset, arg=self.on_finish_arg)
				self.clear()
			queueHandler.queueFunction(queueHandler.eventQueue, h)