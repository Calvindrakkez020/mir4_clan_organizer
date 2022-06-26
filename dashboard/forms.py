from .models import Clan
from django import apps, forms

class ClanForm(forms.Form):
    clan = forms.ModelChoiceField(queryset=Clan.objects.values_list('name', flat=True), label="Clan")