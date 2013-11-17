from elementtree import ElementTree as ET
import os
import sys

if __name__ == "__main__":
	# set your django setting module here
	os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tres.settings") 
	from django.contrib.auth.models import User
	from userProfile.models import UserProfile

	with open('xls/users.xml', 'rt') as f:
		print "parsing list of users..."
		tree = ET.parse(f)
		root = tree.getroot()
		for user in root.findall('user'):
			first_name = user.find('first_name').text
			last_name = user.find('last_name').text
			email = user.find('email').text
			username = user.find('username').text
			password = user.find('password').text
			gender = user.find('gender').text
			description = user.find('description').text
			location = user.find('location').text
			birthday = user.find('birthday').text
			user_list = User.objects.filter(username=username)
			if len(user_list) == 0:
				if not password:
					password = User.objects.make_random_password(5)
					user.find('password').text = password
				new_user = User.objects.create_user(username, email, password)
				print "Created profile of ", username
			else:
				new_user = user_list[0]
				print "Updated profile of ", username

			new_user.first_name = first_name
			new_user.last_name = last_name
			new_user.save()
			user_profile = UserProfile.objects.get(user=new_user)
			user_profile.gender = gender.lower()
			user_profile.description = description
			user_profile.location = location
			user_profile.save()

		print "Updating users in the xml..."
		tree.write('xls/users.xml')
		print "\nDone...\n"





