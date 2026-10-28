'use strict';
// The native menus in the interface languages (the pages have their own dictionaries in
// src/renderer/i18n). English is the key.
const LABELS = {
  es: {
    File: 'Archivo', 'New Document': 'Nuevo documento', 'New Spreadsheet': 'Nueva hoja de cálculo', 'New Presentation': 'Nueva presentación', 'Isadek Notes': 'Isadek Notes',
    New: 'Nuevo', 'Open…': 'Abrir…', Save: 'Guardar', 'Save As…': 'Guardar como…', 'Export as PDF…': 'Exportar como PDF…', 'Print…': 'Imprimir…', Home: 'Inicio',
    Quit: 'Salir', Close: 'Cerrar', Edit: 'Edición', Undo: 'Deshacer', Redo: 'Rehacer', Cut: 'Cortar', Copy: 'Copiar', Paste: 'Pegar', 'Select All': 'Seleccionar todo',
    'Find and Replace…': 'Buscar y reemplazar…', View: 'Vista', 'Zoom In': 'Acercar', 'Zoom Out': 'Alejar', 'Actual Size': 'Tamaño real', 'Toggle Full Screen': 'Pantalla completa',
    'Toggle Developer Tools': 'Herramientas de desarrollo', Window: 'Ventana', Minimize: 'Minimizar', Help: 'Ayuda', 'Check for Updates…': 'Buscar actualizaciones…',
    'Update Automatically': 'Actualizar automáticamente', 'About Isadek Office': 'Acerca de Isadek Office', 'Help and Support': 'Ayuda y soporte', 'Send Crash Reports': 'Enviar informes de errores', 'Show Crash Reports': 'Mostrar informes de errores', 'Privacy': 'Privacidad'
  },
  fr: {
    File: 'Fichier', 'New Document': 'Nouveau document', 'New Spreadsheet': 'Nouveau classeur', 'New Presentation': 'Nouvelle présentation', 'Isadek Notes': 'Isadek Notes',
    New: 'Nouveau', 'Open…': 'Ouvrir…', Save: 'Enregistrer', 'Save As…': 'Enregistrer sous…', 'Export as PDF…': 'Exporter en PDF…', 'Print…': 'Imprimer…', Home: 'Accueil',
    Quit: 'Quitter', Close: 'Fermer', Edit: 'Édition', Undo: 'Annuler', Redo: 'Rétablir', Cut: 'Couper', Copy: 'Copier', Paste: 'Coller', 'Select All': 'Tout sélectionner',
    'Find and Replace…': 'Rechercher et remplacer…', View: 'Affichage', 'Zoom In': 'Zoom avant', 'Zoom Out': 'Zoom arrière', 'Actual Size': 'Taille réelle', 'Toggle Full Screen': 'Plein écran',
    'Toggle Developer Tools': 'Outils de développement', Window: 'Fenêtre', Minimize: 'Réduire', Help: 'Aide', 'Check for Updates…': 'Rechercher des mises à jour…',
    'Update Automatically': 'Mettre à jour automatiquement', 'About Isadek Office': 'À propos d’Isadek Office', 'Help and Support': 'Aide et support', 'Send Crash Reports': 'Envoyer les rapports de plantage', 'Show Crash Reports': 'Afficher les rapports de plantage', 'Privacy': 'Confidentialité'
  },
  de: {
    File: 'Datei', 'New Document': 'Neues Dokument', 'New Spreadsheet': 'Neue Tabelle', 'New Presentation': 'Neue Präsentation', 'Isadek Notes': 'Isadek Notes',
    New: 'Neu', 'Open…': 'Öffnen…', Save: 'Speichern', 'Save As…': 'Speichern unter…', 'Export as PDF…': 'Als PDF exportieren…', 'Print…': 'Drucken…', Home: 'Startseite',
    Quit: 'Beenden', Close: 'Schließen', Edit: 'Bearbeiten', Undo: 'Rückgängig', Redo: 'Wiederholen', Cut: 'Ausschneiden', Copy: 'Kopieren', Paste: 'Einfügen', 'Select All': 'Alles auswählen',
    'Find and Replace…': 'Suchen und Ersetzen…', View: 'Ansicht', 'Zoom In': 'Vergrößern', 'Zoom Out': 'Verkleinern', 'Actual Size': 'Originalgröße', 'Toggle Full Screen': 'Vollbild',
    'Toggle Developer Tools': 'Entwicklertools', Window: 'Fenster', Minimize: 'Minimieren', Help: 'Hilfe', 'Check for Updates…': 'Nach Updates suchen…',
    'Update Automatically': 'Automatisch aktualisieren', 'About Isadek Office': 'Über Isadek Office', 'Help and Support': 'Hilfe und Support', 'Send Crash Reports': 'Absturzberichte senden', 'Show Crash Reports': 'Absturzberichte anzeigen', 'Privacy': 'Datenschutz'
  },
  pt: {
    File: 'Arquivo', 'New Document': 'Novo documento', 'New Spreadsheet': 'Nova planilha', 'New Presentation': 'Nova apresentação', 'Isadek Notes': 'Isadek Notes',
    New: 'Novo', 'Open…': 'Abrir…', Save: 'Salvar', 'Save As…': 'Salvar como…', 'Export as PDF…': 'Exportar como PDF…', 'Print…': 'Imprimir…', Home: 'Início',
    Quit: 'Sair', Close: 'Fechar', Edit: 'Editar', Undo: 'Desfazer', Redo: 'Refazer', Cut: 'Recortar', Copy: 'Copiar', Paste: 'Colar', 'Select All': 'Selecionar tudo',
    'Find and Replace…': 'Localizar e substituir…', View: 'Exibir', 'Zoom In': 'Ampliar', 'Zoom Out': 'Reduzir', 'Actual Size': 'Tamanho real', 'Toggle Full Screen': 'Tela inteira',
    'Toggle Developer Tools': 'Ferramentas de desenvolvedor', Window: 'Janela', Minimize: 'Minimizar', Help: 'Ajuda', 'Check for Updates…': 'Verificar atualizações…',
    'Update Automatically': 'Atualizar automaticamente', 'About Isadek Office': 'Sobre o Isadek Office', 'Help and Support': 'Ajuda e suporte', 'Send Crash Reports': 'Enviar relatórios de falhas', 'Show Crash Reports': 'Mostrar relatórios de falhas', 'Privacy': 'Privacidade'
  }
};
const LANGS = ['en', 'es', 'fr', 'de', 'pt'];

/** A translator for the menus in one language (English for an unknown one). */
function menuText(lang) {
  const d = LABELS[lang] || {};
  return s => d[s] || s;
}

module.exports = { LABELS, LANGS, menuText };
