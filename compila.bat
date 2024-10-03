@echo on

pyinstaller --onefile EmailContasReceber.py
pyinstaller --onefile EmailSaldoClienteFornecedor.py
pyinstaller --onefile EmailPedidosEmAberto.py