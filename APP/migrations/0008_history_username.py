# Generated by Django 4.1.3 on 2023-01-09 08:47

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('APP', '0007_userdetail'),
    ]

    operations = [
        migrations.AddField(
            model_name='history',
            name='username',
            field=models.CharField(default=1, max_length=100),
            preserve_default=False,
        ),
    ]