# Preenchimento--de-Formulario

🚀 Automação de Abertura de Chamados via ServiceNow com Selenium

Este projeto automatiza o processo de abertura de chamados em um portal ServiceNow, utilizando dados extraídos de uma planilha Excel. A solução foi desenvolvida em Python com foco em agilizar o processo.
<hr>

💡 O que o script faz?  
 <ul>
  <li> Lê uma planilha (automacao.xlsx) contendo os dados das solicitações </li>
  <li> Identifica automaticamente o cabeçalho da planilha </li>
  <li> Preenche formulários web de forma automatizada usando Selenium </li>
  <li> Interage com campos dinâmicos (Select2, inputs e referências) </li>
  <li> Aplica estratégias inteligentes para seleção de opções (match exato, parcial e fallback) </li>
  <li> Envia o formulário e captura o número do chamado gerado (RITM/REQ/etc.) </li>
  <li> Atualiza a planilha com os protocolos gerados </li>
  <li> Continua execuções interrompidas sem duplicar chamados </li>
 </ul>
