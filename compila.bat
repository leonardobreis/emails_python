@echo on

pyinstaller --onefile EmailContasReceber.py
pyinstaller --onefile EmailSaldoClienteFornecedor.py
pyinstaller --onefile EmailPedidosEmAberto.py
pyinstaller --onefile EmailExpedicao.py

pyinstaller --onefile EmailTerceirosEmAberto.py

pyinstaller --onefile EmailNFsRecebimento.py