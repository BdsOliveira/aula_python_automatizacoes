from django.shortcuts import render

def index(request):
    return render(request, 'automatizacoes/index.html')

def executar(request):
    return render(request, 'automatizacoes/sucesso.html')
