from django.urls import path
from cv_app import views

urlpatterns = [
    path('', views.upload_cv, name='upload_cv'),
]
