import * as vscode from 'vscode';
import { exec } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';

export function activate(context: vscode.ExtensionContext) {
  let disposable = vscode.commands.registerCommand('latex-pandocs.exportToWord', () => {
    const editor = vscode.window.activeTextEditor;
    if (editor) {
      const document = editor.document;
      const filePath = document.fileName;

      // Verificar si el archivo tiene la extensión .tex
      if (path.extname(filePath) !== '.tex') {
        vscode.window.showErrorMessage('The current file is not a .tex file.');
        return;
      }

      const outputFilePath = filePath.replace(/\.tex$/, '.docx');
      const bibFilePath = filePath.replace(/\.tex$/, '.bib'); // Asume que el archivo .bib tiene el mismo nombre que el archivo .tex

      // Verificar si el archivo .bib existe
      fs.access(bibFilePath, fs.constants.F_OK, (err) => {
        let pandocCommand = `pandoc "${filePath}" -o "${outputFilePath}"`;

        if (!err) {
          // Si el archivo .bib existe, agregar el argumento --bibliography
          pandocCommand = `pandoc "${filePath}" --bibliography="${bibFilePath}" --csl=chicago-author-date.csl -o "${outputFilePath}"`;
        }

        exec(pandocCommand, (error, stdout, stderr) => {
          if (error) {
            vscode.window.showErrorMessage(`Error: ${stderr}`);
          } else {
            vscode.window.showInformationMessage(`File exported to ${outputFilePath}`);
          }
        });
      });
    }
  });

  context.subscriptions.push(disposable);
}

export function deactivate() {}