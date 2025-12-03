from django.db import models
import random
import string

def generate_code():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))

class Guest(models.Model):
    full_name = models.CharField(max_length=255)
    phone_number = models.CharField(max_length=20)
    email = models.EmailField()
    qr_code_value = models.CharField(max_length=8, unique=True, default=generate_code)
    qr_image = models.ImageField(upload_to='qr_codes/', blank=True, null=True)

    def __str__(self):
        return self.full_name
