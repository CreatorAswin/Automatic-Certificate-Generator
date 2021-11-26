from django.db import models
from django.contrib.auth.models import User
from django.utils.text import slugify

class Event(models.Model):
	user = models.ForeignKey(User, on_delete=models.CASCADE)
	event_name = models.CharField(max_length=250)
	date = models.DateField(auto_now_add=True)
	csv_file = models.FileField(upload_to="certificates/csv_files/")
	template = models.FileField(upload_to="certificates/templates/")
	email_column = models.CharField(max_length=250, null=True, blank=True)
	subject = models.CharField(max_length=250, null=True)
	message = models.TextField(null=True, blank=True)
	slug = models.SlugField(null=True, blank=True)

	def save(self, *args, **kwargs):
		self.slug = slugify(self.event_name)
		super(Event, self).save(*args, **kwargs)

class Participant(models.Model):
	event = models.ForeignKey(Event, on_delete=models.CASCADE)
	email = models.CharField(max_length=250)
	status = models.BooleanField(default=False)