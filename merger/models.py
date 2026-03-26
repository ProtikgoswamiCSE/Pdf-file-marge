from django.db import models
import os

class MergeJob(models.Model):
    MERGE_TYPES = [
        ('PDF', 'PDF'),
        ('CSV', 'CSV'),
        ('EXCEL', 'Excel'),
    ]
    STATUS_CHOICES = [
        ('PENDING', 'Pending'),
        ('COMPLETED', 'Completed'),
        ('FAILED', 'Failed'),
    ]
    name = models.CharField(max_length=255)
    merge_type = models.CharField(max_length=10, choices=MERGE_TYPES)
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='PENDING')
    merged_file = models.FileField(upload_to='merged/', null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.name} ({self.merge_type}) - {self.status}"

class UploadedFile(models.Model):
    job = models.ForeignKey(MergeJob, related_name='uploaded_files', on_delete=models.CASCADE)
    file = models.FileField(upload_to='uploads/%Y/%m/%d/')
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def filename(self):
        return os.path.basename(self.file.name)
