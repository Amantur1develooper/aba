from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect
from django.conf import settings
from .models import Profile


@login_required
def profile_view(request):
    profile, _ = Profile.objects.get_or_create(user=request.user)
    token = profile.get_or_create_token()

    bot_username = getattr(settings, "TELEGRAM_BOT_USERNAME", "")
    if bot_username:
        link_url = f"https://t.me/{bot_username}?start={token}"
    else:
        link_url = None

    if request.method == "POST" and "unlink_tg" in request.POST:
        profile.tg_chat_id = ""
        profile.tg_link_token = ""
        profile.save(update_fields=["tg_chat_id", "tg_link_token"])
        return redirect("accounts:profile")

    return render(request, "registration/profile.html", {
        "profile": profile,
        "token": token,
        "link_url": link_url,
    })
