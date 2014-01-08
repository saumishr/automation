import os
import sys
import xlrd
import xlwt
from os import walk
import gc

if __name__ == "__main__":
	# set your django setting module here
	os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tres.settings") 
	from django.contrib.auth.models import User
	from userProfile.models import UserProfile
	from mezzanine.blog.models import BlogPost, BlogParentCategory, BlogCategory
	from django.template.defaultfilters import slugify
	from django.contrib.auth.models import Group
	import uuid
	from django.conf import settings
	from django.core.files import File
	from actstream import actions
	from mezzanine.generic.models import AssignedKeyword, Keyword
	from mezzanine.core.models import CONTENT_STATUS_PUBLISHED

	MEDIA_URL = "static/media/"

	# Absolute filesystem path to the directory that will hold user-uploaded files.
	# Example: "/home/media/media.lawrence.com/media/"
	MEDIA_ROOT = os.path.join(settings.PROJECT_ROOT, *MEDIA_URL.strip("/").split("/"))

	"""
	Global Configurations.
	Please do not change these indexes. These are strictly as per the xls.
	"""
	XLS_CONTAINER 				= 'assets/xls_processed/'
	LOGOS_CONTAINER 			= 'assets/logos/'
	PASSWORD_FILE_PATH			= 'passwords.xls'
	WEBSITE_URL_INDEX 			= 1
	WEBSITE_NAME_INDEX 			= 2
	EMAIL_INDEX 				= 3
	LOGO_FILE_NAME_INDEX 		= 4
	TAG_INDEX					= 7
	WISHRADIO_CATEGORY_INDEX 	= 8
	STORE_DESCRIPTION_INDEX 	= 10
	MAX_PASSWORD_CHARACTERS		= 5

	store_count = 0

	# def save_file_s3(file, path=''):
	# 	''' Little helper to save a file
	# 	'''
	# 	filename = file._get_name()
	# 	new_filename =  u'{name}.{ext}'.format(  name=uuid.uuid4().hex,
	# 											ext=os.path.splitext(filename)[1].strip('.'))

	# 	dir_path =  str(path)

	# 	save_path = os.path.join(dir_path, new_filename)
	# 	storage=S3BotoStorage(location=settings.STORAGE_ROOT)
	# 	storage.save(save_path, file)

	# 	return save_path

	def save_file(file, path=''):
		''' Little helper to save a file
		'''
		filename = file.name
		# Enable this code to save the file with uid.
		# new_filename =  u'{name}.{ext}'.format(	name=uuid.uuid4().hex,
		# 										ext=os.path.splitext(filename)[1].strip('.'))
		# else
		new_filename =  os.path.basename(filename)

		dir_path =  '%s/%s' % (MEDIA_ROOT, str(path))

		if not os.path.exists(dir_path):
			os.makedirs(dir_path)

		save_path = os.path.join(dir_path, new_filename)
		print "save path: ", save_path
		print "filename: ", new_filename
		with open(save_path, 'wb+') as destination:
			for chunk in file.chunks():
				destination.write(chunk)
			destination.close()

		return os.path.join(path, new_filename)

	def create_store(sheet, start_index, end_index, store_count, current_tag):
		"""
			Single item cloumns will have store details which are unique to them and will be used to create stores.
			Also, these entries as expected to be in the first row of their respective row-range. 
		"""
		store_details 	= sheet.row_values(start_index)

		url 			= store_details[WEBSITE_URL_INDEX]
		store_name 		= store_details[WEBSITE_NAME_INDEX]
		email 			= store_details[EMAIL_INDEX]
		description 	= store_details[STORE_DESCRIPTION_INDEX]
		logo_name 		= store_details[LOGO_FILE_NAME_INDEX]
		tags			= store_details[TAG_INDEX]

		print '------------------------------------------------------------------------------------------'
		print 'Store: ', store_name
		print '------------------------------------------------------------------------------------------'
		print 'fetching store details...'
		if logo_name:
			"""
				extract the file name & extension.
			"""
			logo 	= os.path.splitext(logo_name)[0]
			ext 	= os.path.splitext(logo_name)[1]


			"""
				Now start the dirty work. :(
				Our contractor totally messed up the file names. This code is a corrective measure to clear up the mess.
				This code is totally redudant once the logo names in the xls and logo file name are consistent.
			"""

			"""
				First issue: extra .png between filename & extension.
			"""
			orig_logo_name = str(current_tag)+'.' + logo + '.png' + ext


			"""
				Expected logo name.
			"""
			logo_name = str(current_tag)+'.' + logo + ext.lower()
			path = logoFolder +  '/%s/%s' % (sheet.name, orig_logo_name)


			"""
				If path with extra .png in filename exists, remove the extra .png .
			"""
			if os.path.exists(path):
				newpath = logoFolder +  '/%s/%s' % (sheet.name, logo_name) 
				os.rename(path, newpath)
			else:
				"""
					Second issue: ext name is in uppercase at some places.
				"""
				orig_logo_name = str(current_tag)+'.' + logo + '.png' + ext.upper()
				path = logoFolder +  '/%s/%s' % (sheet.name, orig_logo_name)
				if os.path.exists(path):
					newpath = logoFolder +  '/%s/%s' % (sheet.name, logo_name)
					os.rename(path, newpath)
			"""
				Rest all the issues are handled manually as they are not genric enough to be handled in automated script.
			"""

		wr_categories = sheet.col_values(WISHRADIO_CATEGORY_INDEX, start_rowx=start_index, end_rowx=end_index)
		wr_categories = filter(None, wr_categories)
		print 'categories: ', wr_categories

		"""
			users will be created with default username: store_name + '_admin'
		"""
		if len(store_name) > 24:
			store_name = store_name[0:23]

		username = store_name + '_admin'
		user_list = User.objects.filter(username=username)
		new_user = None

		if len(user_list) == 0:
			print 'Creating profile for admin of store: ', store_name
			password = User.objects.make_random_password(MAX_PASSWORD_CHARACTERS)
			email = email.split()			
			if len(email) > 0:
				email = email[0]
			else:
				email = ''

			print "username:", username, " email:", email, " password: ", password
			new_user = User.objects.create_user(username, email, password)
			new_user.first_name = username
			
			new_user.is_staff = True
			group = Group.objects.get(name='StoreOwners') 
			group.user_set.add(new_user)

			new_user.save()
			user_profile = UserProfile.objects.get(user=new_user)
			user_profile.gender = 'male'
			user_profile.save()
			passsheet.write(store_count, 0, username)
			passsheet.write(store_count, 1, password)
			print "Created profile of ", username
		else:
			print "Admin for store: ", store_name, " already exists..."
			new_user = user_list[0]

		print 'Creating store: ', store_name
		blog_post = None
		blog_post_list = BlogPost.objects.filter(user=new_user)
		if len(blog_post_list) == 0:
			blog_post = BlogPost.objects.create(web_url=url, content=description, title=store_name, user=new_user, status=CONTENT_STATUS_PUBLISHED)
			blog_post.save()
		else:
			print 'Store ', store_name, ' already exists...'
			blog_post = blog_post_list[0]

		print 'Adding categories...'
		"""
		ToDo: Need to check the existing list of categories with cateories mentioned in the xls.
		Removed categories should also be updated in the blog_post object.
		"""
		for wr_parent_category in wr_categories:
			parent_category_list = BlogParentCategory.objects.filter(title=wr_parent_category)
			if len(parent_category_list) != 0:
				parent_category = parent_category_list[0]
				sub_categories = BlogCategory.objects.all().filter(parent_category=parent_category)
				for sub_category in sub_categories:
					if not blog_post.categories.all().filter(slug=slugify(sub_category)).exists():
						blog_post.categories.add(sub_category)

		print 'Attaching logo with the store...'				
		if logo_name:
			new_file_rel_path = 'users/store/%s/images/' % (blog_post.id)
			path = logoFolder +  '/%s/%s' % (sheet.name, logo_name)
			"""
				Remove any spaces in the path.
			"""
			path = path.replace(" ", "")

			featuredImageObj =  File(open(path, 'r'))
			new_file_path = save_file(featuredImageObj, new_file_rel_path)
			if blog_post.featured_image:
				old_file_path = '%s/%s' % (MEDIA_ROOT, str(blog_post.featured_image.path))
				if os.path.exists(old_file_path):
					print 'Removed old logo...'
					os.remove(old_file_path)

			print 'new logo is set...:', new_file_path
			blog_post.featured_image = new_file_path

		blog_post.save()
		blogpost_tags = sheet.col_values(TAG_INDEX, start_rowx=start_index, end_rowx=end_index)
		blogpost_tags = filter(None, blogpost_tags)

		for kw in blogpost_tags:
			kw = kw.strip().lower()
			if kw:
				keyword_id = Keyword.objects.get_or_create(title=kw)[0].id
				blog_post.keywords.add(AssignedKeyword(keyword_id=keyword_id))

		blog_post = BlogPost.objects.get(id=blog_post.id)
		if blog_post and new_user:
			actions.follow(new_user, blog_post, send_action=False, actor_only=False) 
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

		logoFolder = LOGOS_CONTAINER + os.path.splitext(filename)[0]
		"""
			'assets/xls' is supposed to contain all the xls files having store information.
		"""
		workbook = xlrd.open_workbook(XLS_CONTAINER + filename)

		"""
			passbook is the xls file generated in the end. This contains all the login credentials of the store owners.
		"""
		passbook = xlwt.Workbook()
		passsheet = passbook.add_sheet('passwords')
		passsheet.write(0, 0, 'username')
		passsheet.write(0, 1, 'password')

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
					create_store(sheet, start_index, end_index, store_count, start_tag )
					start_tag = end_tag
					end_tag = start_tag + 1
				else:
					create_store(sheet, start_index, sheet.nrows, store_count, start_tag)
					break

	passbook.save(PASSWORD_FILE_PATH)


		






