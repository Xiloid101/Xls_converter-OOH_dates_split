from django.db import models


class UpFile(models.Model):
    title = models.CharField(max_length=50)
    upload = models.FileField(upload_to='media/')

