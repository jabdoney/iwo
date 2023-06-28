"""iwo URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from .views import home,input,download,readword,help,import_error
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('bar-services/iwo/home/',home),
    path('bar-services/iwo/download/',download),
    path('bar-services/iwo/help/',help),
    path('bar-services/iwo/import-error/',import_error),
    path('bar-services/iwo/input/',input,name="input"),
    path('bar-services/iwo/readword/',readword),
    path('bar-services/iwo/admin/', admin.site.urls,name='root')
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)