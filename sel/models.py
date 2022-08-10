from django.db import models

# Create your models here.
class mg(models.Model):
    name=models.CharField(max_length=10)
    email=models.EmailField(max_length=20)
    feed=models.CharField(max_length=100)
