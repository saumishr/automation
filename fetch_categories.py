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
	from django.conf import settings

	MEDIA_URL = "static/media/"

	# Absolute filesystem path to the directory that will hold user-uploaded files.
	# Example: "/home/media/media.lawrence.com/media/"
	MEDIA_ROOT = os.path.join(settings.PROJECT_ROOT, *MEDIA_URL.strip("/").split("/"))

	"""
	Global Configurations.
	Please do not change these indexes. These are strictly as per the xls.
	"""
	CATEGORIES_FILE_PATH			= '/home/saumishr/passcodes/wishradio_categories.xls'

	def fetch_category_relations():
		store_count = 0
		serial_pointer = 1

		blog_posts = BlogPost.objects.all()

		for blog_post in blog_posts:
			store_count += 1
			print 'Processing Store: ', blog_post.title
			passsheet.write(serial_pointer, 0, store_count)
			passsheet.write(serial_pointer, 1, blog_post.title)
			parent_categories = BlogParentCategory.objects.all()
			hasCategories = False
			for parent_category in parent_categories:
				sub_categories_list = []
				blog_post_added_sub_categories = blog_post.categories.all()
				sub_categories = BlogCategory.objects.all().filter(parent_category=parent_category)
				for sub_category in sub_categories:
					if sub_category in blog_post_added_sub_categories:
						sub_categories_list.append(sub_category.title)

				if len(sub_categories_list) > 0:
					passsheet.write(serial_pointer, 2, parent_category.title)
					for sub_category in sub_categories_list:
						passsheet.write(serial_pointer, 3, sub_category)
						serial_pointer += 1
						hasCategories = True

			if hasCategories == False:
				passsheet.write(serial_pointer, 2, '')
				passsheet.write(serial_pointer, 3, '')
				serial_pointer += 1	

		gc.collect()


	"""
		Main()
	"""
	passbook = xlwt.Workbook()
	passsheet = passbook.add_sheet('categories')
	passsheet.write(0, 0, 'Serial')
	passsheet.write(0, 1, 'Store')
	passsheet.write(0, 2, 'Parent Category')
	passsheet.write(0, 3, 'Sub Category')

	fetch_category_relations()
	

	passbook.save(CATEGORIES_FILE_PATH)


		






