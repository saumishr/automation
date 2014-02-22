import os

if __name__ == "__main__":
	# set your django setting module here
	os.environ.setdefault("DJANGO_SETTINGS_MODULE", "tres.settings") 
	from mezzanine.blog.models import BlogPost, BlogParentCategory, BlogCategory
	from mezzanine.generic.models import AssignedKeyword, Keyword

	def tag_stores():
		blog_posts = BlogPost.objects.all()
		for blog_post in blog_posts:
			title = blog_post.title
			keyword_id = Keyword.objects.get_or_create(title=title)[0].id
			blog_post.keywords.add(AssignedKeyword(keyword_id=keyword_id))
			print "Added keyword: ", title

	"""
		Main()
	"""

	tag_stores()


		






