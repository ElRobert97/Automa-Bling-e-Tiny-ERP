import multiprocessing
import os

jf_miguel = r'C:\Users\Contler Elias\OneDrive - AGORA DEU LUCRO EDUCACIONAL LTDA\Área de Trabalho\Hashtag\Bots\Bling ERP\JF Miguel\bling_erp.py'
proper4 = r'C:\Users\Contler Elias\OneDrive - AGORA DEU LUCRO EDUCACIONAL LTDA\Área de Trabalho\Hashtag\Bots\Bling ERP\Proper4\bling_erp.py'
focco_ecom =r'C:\Users\Contler Elias\OneDrive - AGORA DEU LUCRO EDUCACIONAL LTDA\Área de Trabalho\Hashtag\Bots\Bling ERP\Focco\tinyselenium.py'
albha = r'C:\Users\Contler Elias\OneDrive - AGORA DEU LUCRO EDUCACIONAL LTDA\Área de Trabalho\Hashtag\Bots\Bling ERP\Albha\bling_erp.py'

python_executable = r'C:\Users\Contler Elias\AppData\Local\Programs\Python\Python312\Python.exe'

# Função para executar cada script
def run_script(script):
    os.system(f'python "{script}"')


if __name__ == '__main__':
    # Lista de scripts a serem executados
    scripts =  [jf_miguel, albha, focco_ecom]

    # Criar e iniciar um processo para cada script
    processes = []
    for script in scripts:
            process = multiprocessing.Process(target=run_script, args=(script,))
            process.start()
            processes.append(process)

    # Aguardar todos os processos terminarem
    for process in processes:
        process.join()