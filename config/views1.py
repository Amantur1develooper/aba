from django.shortcuts import redirect


def home(request):
    if request.user.is_authenticated:
        return redirect("dds:hotel_list")
    return redirect("/accounts/login/")