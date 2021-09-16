from django.core import mail
from django.http import response
from django.http.request import HttpRequest
from django.shortcuts import render, redirect
from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from .forms import SignupForm
from django.contrib.sites.shortcuts import get_current_site
from django.utils.encoding import force_bytes, force_text
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from .decorators import unauthenticated_user, allowed_users
from django.contrib.auth.models import Group
from openpyxl import load_workbook
import os
import shutil
from oletools.olevba3 import VBA_Parser
import re
from django.utils.http import urlsafe_base64_encode, urlsafe_base64_decode
from django.template.loader import render_to_string
from .tokens import account_activation_token
from django.contrib.auth.models import User
from django.core.mail import EmailMessage, message


@login_required(login_url='/registration/login/')
def login_request(request):
    if request.method == 'POST':
        form = AuthenticationForm(request=request, data=request.POST)
        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)
            if user is not None:
                login(request, user)
                messages.info(request, f"You are now logged in as {username}")
            return render(request, "upload.html")

    form = AuthenticationForm()
    return render(request = request,
                    template_name = "login.html",
                    context={"form":form})

@unauthenticated_user
def login(request):
    """
    Creates login view
    Returns: rendered login page
    """
    return render(request, 'login.html')


def signup(request):
    if request.method == 'POST':
        form = SignupForm(request.POST)
        if form.is_valid():
            user = form.save(commit=False)
            user.is_active = False
            user.save()
            current_site = get_current_site(request)
            mail_subject = 'Activite your account.'
            message = render_to_string('acc_active_email.html',{
                'user': user,
                'domain': current_site.domain,
                'uid': urlsafe_base64_encode(force_bytes(user.pk)),
                'token': account_activation_token.make_token(user),
            })
            to_email = form.cleaned_data.get('email')
            email = EmailMessage(
                mail_subject, message, to=[to_email]
            )
            email.send()    
            if request.POST.get('group') == 'current':
                group = Group.objects.get(name='New user')
            user.groups.add(group)
            return response.HttpResponse('Please confirm your email addres to complete the registration')
    else:
        form = SignupForm()
    return render(request, 'signup.html', {'form': form})

def logout_request(request):
    logout(request)
    messages.info(request, "you have successfully logged out.")
    return redirect('login')

def activate(request, uidb64, token):
    try:
        uid = force_text(urlsafe_base64_decode(uidb64))
        user = User.objects.get(pk=uid)
    except(TypeError, ValueError, OverflowError, User.DoesNotExist):
        user = None
    if user is not None and account_activation_token.check_token(user, token):
        user.is_active = True
        user.save()
        login(request, user)
        # return redirect('home')
        return response.HttpResponse('Thank you for your email confirmation. Now you can login your account.')
    else:
        return response.HttpResponse('Activation link is invalid!')

@allowed_users(allowed_roles=["creator"])
def upload(request):
    if request.method == 'POST':
        if request.FILES.get('document'):
            file = request.FILES['document']

            EXCEL_FILE_EXTENSIONS = ('xlsb', 'xls', 'xlsm', 'xla', 'xlt', 'xlam',)
            KEEP_NAME = False  # Set this to True if you would like to keep "Attribute VB_Name"
            
            workbook = load_workbook(filename=file, data_only=True)
            xls = workbook[workbook.sheetnames[0]] 
            
            result = []
            # macros check
            vba_parser = VBA_Parser(file)
            vba_modules = vba_parser.extract_all_macros()
            if vba_parser.detect_vba_macros():
                result.append("Caution, macros has been found in your excel file.")


            count = 0
            
            # check .exe, url 
            for i in range(1, xls.max_row+1):
                count = 0
                for cell in xls[i]:
                    for column_value in cell.value.split(','):
                        count = count + 1
                        if re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', str(column_value)):
                            message = str('found a url link in  row ', i ,' and column ', count)
                            result.append(message)

                        elif re.findall(r'\b\S*\.exe\b', str(column_value)):
                            message = str( 'found .exe file in  row ', i ,' and column ', count)
                            result.append(message)

                        elif count == len(xls[i]):
                            count = 0

            
            # check hidden rows
            for rowLetter,rowDimension in xls.row_dimensions.items():
                if rowDimension.hidden == True:
                    hidden_row = str('found hidden rows at ', rowLetter)
                    result.append(hidden_row)


            dic_result = { 'report': result}
            return render(request, 'upload.html', dic_result)
    return render(request, 'upload.html')