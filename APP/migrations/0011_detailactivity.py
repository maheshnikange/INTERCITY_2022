# Generated by Django 4.1.3 on 2023-01-11 03:58

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('APP', '0010_delete_userdetail'),
    ]

    operations = [
        migrations.CreateModel(
            name='DetailActivity',
            fields=[
                ('ID', models.IntegerField(primary_key=True, serialize=False)),
                ('username', models.CharField(max_length=100)),
                ('logidate', models.DateTimeField()),
                ('logodate', models.DateTimeField()),
                ('acti', models.CharField(max_length=100)),
            ],
        ),
    ]
