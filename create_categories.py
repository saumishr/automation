from elementtree import ElementTree as ET
import os
import sys
import xlrd

if __name__ == "__main__":
	# set your django setting module here
	os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tres.settings") 
	from django.contrib.auth.models import User
	from userProfile.models import UserProfile
	from mezzanine.blog.models import BlogParentCategory, BlogCategory
	from django.template.defaultfilters import slugify

	workbook = xlrd.open_workbook('xls/Wishradio - Categories.xlsx')

	def create_categories(sheet, start_index, end_index):
		categories = sheet.row_values(start_index)
		parent_category = None
		parent_category_list = BlogParentCategory.objects.filter(title=categories[1])
		if len(parent_category_list) == 0:
			parent_category = BlogParentCategory(title=categories[1], slug=slugify(categories[1]))
			parent_category.save()
		else:
			parent_category = parent_category_list[0]

		sub_categories = sheet.col_values(2, start_rowx=start_index, end_rowx=end_index)
		for sub_category in sub_categories:
			sub_category_list = BlogCategory.objects.filter(title=sub_category)
			if len(sub_category_list) == 0:
				sub_category = BlogCategory(title=sub_category, parent_category=parent_category, slug=slugify(sub_category))
				sub_category.save()
			else:
				sub_category = sub_category_list[0]
				if sub_category.parent_category.title == parent_category.title:
					print parent_category, " already has a sub category: ", sub_category
				else:
					sub_category.parent_category = parent_category


	for sheet in workbook.sheets():
			column_0_values = sheet.col_values(colx=0)

			start_tag = 1
			end_tag = start_tag + 1
			print "Processing sheet: ", sheet.name

			while(start_tag in column_0_values):
				start_index = column_0_values.index(start_tag)
				if end_tag in column_0_values:
					end_index = column_0_values.index(end_tag)
					start_tag = end_tag
					end_tag = start_tag + 1
					create_categories(sheet, start_index, end_index )
				else:
					create_categories(sheet, start_index, sheet.nrows)
					break

		






