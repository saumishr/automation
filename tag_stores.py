import os
import sys
import xlrd
import xlwt
from os import walk
import gc

if __name__ == "__main__":
	# set your django setting module here
	os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tres.settings") 
	from mezzanine.blog.models import BlogPost
	import uuid
	from django.conf import settings
	from django.core.files import File
	from actstream import actions
	from mezzanine.generic.models import AssignedKeyword, Keyword
	from mezzanine.core.models import CONTENT_STATUS_PUBLISHED
	from django.contrib.contenttypes.models import ContentType

	MEDIA_URL = "static/media/"

	# Absolute filesystem path to the directory that will hold user-uploaded files.
	# Example: "/home/media/media.lawrence.com/media/"
	MEDIA_ROOT = os.path.join(settings.PROJECT_ROOT, *MEDIA_URL.strip("/").split("/"))

	"""
	Global Configurations.
	Please do not change these indexes. These are strictly as per the xls.
	"""
	XLS_CONTAINER 				= 'assets/xls_processed/'
	WEBSITE_NAME_INDEX 			= 2
	TAG_INDEX					= 7

	store_count = 0


	def tag_store(sheet, start_index, end_index, store_count, current_tag):
		"""
			Single item cloumns will have store details which are unique to them and will be used to create stores.
			Also, these entries as expected to be in the first row of their respective row-range. 
		"""
		store_details 	= sheet.row_values(start_index)

		store_name 		= store_details[WEBSITE_NAME_INDEX]

		print '------------------------------------------------------------------------------------------'
		print 'tagging store: ', store_name
		print '------------------------------------------------------------------------------------------'

		blog_post = None
		blog_post = BlogPost.objects.get(title=store_name)
		ctype = ContentType.objects.get_for_model(BlogPost)
		object_pk = blog_post.id

		blogpost_tags = sheet.col_values(TAG_INDEX, start_rowx=start_index, end_rowx=end_index)
		blogpost_tags = filter(None, blogpost_tags)
		blog_post.keywords.clear()
		blog_post		= BlogPost.objects.get(title=store_name)

		for kw in blogpost_tags:
			kw = kw.strip().lower()
			if kw:
				keyword = Keyword.objects.get_or_create(title=kw)[0]
				if keyword:
					assignedKeywords = AssignedKeyword.objects.all().filter(keyword=keyword, content_type=ctype, object_pk=object_pk)
				
					if len(assignedKeywords) == 0:
						assignedKeyword = AssignedKeyword(keyword=keyword)
						blog_post.keywords.add(assignedKeyword)

		sub_categories = blog_post.categories.all()
		for kw in sub_categories:
			kw = kw.title.strip().lower()
			if kw:
				keyword = Keyword.objects.get_or_create(title=kw)[0]
				if keyword:
					assignedKeywords = AssignedKeyword.objects.filter(keyword=keyword, content_type=ctype, object_pk=object_pk)
					if len(assignedKeywords) == 0:
						assignedKeyword = AssignedKeyword(keyword=keyword)
						blog_post.keywords.add(assignedKeyword)

		gc.collect()

	"""
		Main()
	"""
	f = []
	for (dirpath, dirnames, filenames) in walk(XLS_CONTAINER):
		f.extend(filenames)
		break

	print '------------------------------------------------------------------------------------------'
	print "Indexing files : ", f
	print '------------------------------------------------------------------------------------------'

	for filename in f:
		print '------------------------------------------------------------------------------------------'
		print 'started Processing file: ' + filename
		print '------------------------------------------------------------------------------------------'

		"""
			'assets/xls' is supposed to contain all the xls files having store information.
		"""
		workbook = xlrd.open_workbook(XLS_CONTAINER + filename)

		"""
			passbook is the xls file generated in the end. This contains all the login credentials of the store owners.
		"""

		sheets = workbook.sheets()
		
		print '------------------------------------------------------------------------------------------'
		print 'getting sheets in ', filename
		print 'sheets found: ', [sheet.name for sheet in sheets]
		print '------------------------------------------------------------------------------------------'

		for sheet in sheets:
			"""
			column 0 of every sheet is supposed to containe indexes of the stores.
			"""
			column_0_values = sheet.col_values(colx=0)

			"""
			column 0 indexes should always start from 1 and not 0. Also these indexes should always be in the first row of their row-range.
			To fetch store data from sheets we follow following logic:
			* Get Row index of start_tag
			* Get Row index of end_tag. (end_tag = start_tag + 1)
			* All the data between these rows will be of current store.
			"""
			start_tag = 1 
			end_tag = start_tag + 1

			while(start_tag in column_0_values):
				store_count = store_count + 1
				start_index = column_0_values.index(start_tag)
				if end_tag in column_0_values:
					end_index = column_0_values.index(end_tag)
					"""
						This API will fetch all the rows between rang(start_index, end_index) and process the data.
						Will create users and their respective stores.
					"""
					tag_store(sheet, start_index, end_index, store_count, start_tag )
					start_tag = end_tag
					end_tag = start_tag + 1
				else:
					tag_store(sheet, start_index, sheet.nrows, store_count, start_tag)
					break


		






