from django.shortcuts import render, redirect
from django.http import HttpResponse
from .models import Character, Clan

# Create your views here.
def dashboard(request):
    clans_queryset = Clan.objects.all()
    characters_queryset = Character.objects.all().order_by('-level')

    content = {
        "clans": clans_queryset,
        "characters": characters_queryset
    }

    return render(request, "index.html", content)

def index(request):
    redirect('/dashboard')
