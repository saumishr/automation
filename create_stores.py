from elementtree import ElementTree as ET
import os
import sys
import xlrd
import xlwt

if __name__ == "__main__":
	# set your django setting module here
	os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tres.settings") 
	from django.contrib.auth.models import User
	from userProfile.models import UserProfile
	from mezzanine.blog.models import BlogPost, BlogParentCategory, BlogCategory

	workbook = xlrd.open_workbook('xls/1.Fashion.xls')
	passbook = xlwt.Workbook()
	passsheet = passbook.add_sheet('passwords')
	passsheet.write(0, 0, 'username')
	passsheet.write(0, 1, 'password')


	def create_store(sheet, start_index, end_index, user_count):
		store_details = sheet.row_values(start_index)

		url = store_details[1]
		store_name = store_details[2]
		email = store_details[3]
		description = store_details[10]
		wr_categories = sheet.col_values(8, start_rowx=start_index, end_rowx=end_index)
		wr_categories = filter(None, wr_categories)

		#Create the user first
		username = store_name + '_admin'
		user_list = User.objects.filter(username=username)
		new_user = None
		if len(user_list) == 0:
			password = User.objects.make_random_password(5)
			email = email.split()
			print "username:", username, " email:", email[0], " password: ", password
			new_user = User.objects.create_user(username, email[0], password)
			new_user.first_name = username
			new_user.save()
			user_profile = UserProfile.objects.get(user=new_user)
			user_profile.gender = 'male'
			user_profile.save()
			passsheet.write(user_count, 0, username)
			passsheet.write(user_count, 1, password)
			print "Created profile of ", username
		else:
			print "user ", username, " already exists..."
			new_user = user_list[0]

		blog_post = None
		blog_post_list = BlogPost.objects.filter(user=new_user)
		if len(blog_post_list) == 0:
			blog_post = BlogPost(web_url=url, content=description, title=store_name, user=new_user)
			blog_post.save()
		else:
			blog_post = blog_post_list[0]

		for wr_parent_category in wr_categories:
			parent_category_list = BlogParentCategory.objects.filter(title=wr_parent_category)
			if len(parent_category_list) != 0:
				parent_category = parent_category_list[0]
				sub_categories = BlogCategory.objects.all().filter(parent_category=parent_category)
				for sub_category in sub_categories:
					if not BlogPost.objects.all().filter(categories=sub_category).exists():
						blog_post.categories.add(sub_category)

		blog_post.save()


	user_count = 0
	for sheet in workbook.sheets():
			column_0_values = sheet.col_values(colx=0)

			start_tag = 1
			end_tag = start_tag + 1
			print "Processing sheet: ", sheet.name

			while(start_tag in column_0_values):
				user_count = user_count + 1
				start_index = column_0_values.index(start_tag)
				if end_tag in column_0_values:
					end_index = column_0_values.index(end_tag)
					start_tag = end_tag
					end_tag = start_tag + 1
					create_store(sheet, start_index, end_index, user_count )
				else:
					create_store(sheet, start_index, sheet.nrows, user_count)
					break

	passbook.save('xls/passwords.xls')


		






