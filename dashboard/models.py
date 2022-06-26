from django.db import models
from django.dispatch import receiver

# Create your models here.
class Clan(models.Model):
    name = models.CharField(max_length=50)
    total_members = models.IntegerField(default=0)
    gold_donation = models.IntegerField(default=0)

    def __str__(self):
        return self.name

class Character(models.Model):
    name = models.CharField(max_length=50)
    level = models.IntegerField(default=0)
    gold_donation = models.IntegerField(default=0)
    gold_debt = models.IntegerField(default=0)
    advanced_gold = models.IntegerField(default=0)
    clan = models.ForeignKey(Clan, on_delete=models.CASCADE, null=True)

    def __str__(self):
        return self.name
    
    def save(self, *args, **kwargs):
        try:
            old_char = Character.objects.get(pk=self.pk)
            try:
                if self.clan != old_char:
                    old_char.clan.total_members -= 1
                    old_char.clan.save()
            except AttributeError:
                self.clan.total_members += 1
                self.clan.save()
            super(Character, self).save(*args, **kwargs)
        except Character.DoesNotExist:
            super(Character, self).save(*args, **kwargs)

    def delete(self):
        clan = self.clan
        clan.total_members -= 1
        super(Character, self).delete()

@receiver(models.signals.post_save, sender=Character)
def execute_after_save(sender, instance, created, *args, **kwargs):
    if created:
        clan = instance.clan
        clan.total_members += 1
        clan.save()
