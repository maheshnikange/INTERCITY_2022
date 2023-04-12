# Generated by Django 4.1.3 on 2023-01-11 07:58

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('APP', '0013_remove_user_role'),
    ]

    operations = [
        migrations.AddField(
            model_name='user',
            name='role',
            field=models.CharField(choices=[('SuperAdmin', 'SuperAdmin'), ('Admin', 'Admin'), ('User', 'User')], default=1, max_length=10),
            preserve_default=False,
        ),
    ]