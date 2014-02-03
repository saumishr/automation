import os
import sys
import xlrd
import xlwt
from os import walk
import gc

if __name__ == "__main__":
	# set your django setting module here
	os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tres.settings") 
	from mezzanine.blog.models import BlogPost, BlogParentCategory, BlogCategory
	from mezzanine.generic.models import AssignedKeyword, Keyword
	from django.conf import settings

	MEDIA_URL = "static/media/"

	# Absolute filesystem path to the directory that will hold user-uploaded files.
	# Example: "/home/media/media.lawrence.com/media/"
	MEDIA_ROOT = os.path.join(settings.PROJECT_ROOT, *MEDIA_URL.strip("/").split("/"))

	"""
	Global Configurations.
	Please do not change these indexes. These are strictly as per the xls.
	"""
	CATEGORY_MAPPING_FILE_PATH			= 'assets/category_mapping/Wishradio-Sub-category-mapping.xls'
	MAXIMUM_STORE_INDEX					= 247
	STORE_NAME_INDEX					= 1
	WISHRADIO_CATEGORY_INDEX			= 2
	WISHRADIO_SUBCATEGORY_INDEX			= 3


	store_count = 0

	def create_category_mapping(sheet, start_index, end_index, store_count, current_tag):
		store_details 	= sheet.row_values(start_index)
		store_name 		= store_details[STORE_NAME_INDEX]
		print "mapping store: ",store_count,":", store_name

		blog_post 		= BlogPost.objects.get(title=store_name)
		blog_post.categories.clear()

		sub_categories = sheet.col_values(WISHRADIO_SUBCATEGORY_INDEX, start_rowx=start_index, end_rowx=end_index)
		sub_categories = filter(None, sub_categories)

		for sub_category in sub_categories:
			sub_category_list = BlogCategory.objects.filter(title=sub_category)
			if len(sub_category_list) != 0:
				sub_category = sub_category_list[0]
				blog_post.categories.add(sub_category)

		blog_post.save()

		for kw in sub_categories:
			kw = kw.strip().lower()
			if kw:
				keyword_id = Keyword.objects.get_or_create(title=kw)[0].id
				blog_post.keywords.add(AssignedKeyword(keyword_id=keyword_id))

		gc.collect()


	"""
		Main()
	"""
	workbook = xlrd.open_workbook(CATEGORY_MAPPING_FILE_PATH)

	"""
		passbook is the xls file generated in the end. This contains all the login credentials of the store owners.
	"""

	sheets = workbook.sheets()

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

		while(start_tag in column_0_values): #and start_tag <= MAXIMUM_STORE_INDEX):
			store_count = store_count + 1
			start_index = column_0_values.index(start_tag)

			if end_tag in column_0_values:
				end_index = column_0_values.index(end_tag)
				"""
					This API will fetch all the rows between rang(start_index, end_index) and process the data.
					Will create users and their respective stores.
				"""
				create_category_mapping(sheet, start_index, end_index, store_count, start_tag )
				start_tag = end_tag
				end_tag = start_tag + 1
			else:
				create_category_mapping(sheet, start_index, sheet.nrows, store_count, start_tag)
				break
	


		






