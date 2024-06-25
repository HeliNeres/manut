if (Test-Path "venv") {
    Write-Output "Ambiente virtual 'venv' encontrado."

    .\venv\Scripts\Activate

    Write-Output "Ambiente virtual ativado."
} else {
    Write-Output "Ambiente virtual 'venv' não encontrado. Criando..."

    python -m venv venv

    Write-Output "Ambiente virtual criado."

    .\venv\Scripts\Activate

    Write-Output "Ambiente virtual ativado."

    pip install -r requirements.txt
    python.exe -m pip install --upgrade pip
    Write-Output "Requisitos instalados."
}

python main.py

Read-Host -Prompt "Press Enter to exit."