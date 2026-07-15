/* ============================================================
   APP.JS
   Ponto de entrada. Todos os outros módulos já registraram suas
   constantes, funções e listeners; aqui só disparamos o
   carregamento do estado para iniciar o sistema.

   IMPORTANTE: este arquivo deve ser o ÚLTIMO <script> carregado
   na página, na seguinte ordem:

   config.js -> state.js -> utils.js -> filters.js -> excel.js ->
   sharepoint.js -> report.js -> render.js -> events.js -> app.js
============================================================ */

loadState();
