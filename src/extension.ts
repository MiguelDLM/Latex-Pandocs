import * as vscode from 'vscode';
import { exec } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as mammoth from 'mammoth';

export function activate(context: vscode.ExtensionContext) {
  // Comando para exportar a Word
  let exportToWordDisposable = vscode.commands.registerCommand('latex-pandocs.exportToWord', () => {
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

  // Comando para convertir Word a Tex
  let wordToTexDisposable = vscode.commands.registerCommand('latex-pandocs.wordToTex', async () => {
    const editor = vscode.window.activeTextEditor;
    if (editor) {
      const document = editor.document;
      const filePath = document.fileName;

      // Verificar si el archivo tiene la extensión .tex
      if (path.extname(filePath) !== '.tex') {
        vscode.window.showErrorMessage('The current file is not a .tex file.');
        return;
      }

      const docxFilePath = filePath.replace(/\.tex$/, '.docx');
      const finalTexFilePath = filePath.replace(/\.tex$/, '_reviewed.tex');

      // Verificar si el archivo .docx existe
      fs.access(docxFilePath, fs.constants.F_OK, async (err) => {
        if (err) {
          vscode.window.showErrorMessage(`The corresponding .docx file (${docxFilePath}) does not exist.`);
          return;
        }

        try {
          // Convertir el archivo .docx a HTML usando mammoth con opciones
          const options = {
            styleMap: [
              "comment-reference => sup"
            ]
          };
          const result = await mammoth.convertToHtml({ path: docxFilePath }, options);
          let html = result.value;

          // Procesar el HTML para extraer comentarios y revisiones
          const comments = extractComments(html);

          // Insertar comentarios en el HTML
          comments.forEach(({ comment, position }) => {
            html = html.slice(0, position) + `<!-- ${comment} -->` + html.slice(position);
          });

          // Convertir el HTML a LaTeX usando Pandoc
          const pandocCommand = `pandoc -f html -t latex -o "${finalTexFilePath}"`;
          const pandocProcess = exec(pandocCommand, (error, stdout, stderr) => {
            if (error) {
              vscode.window.showErrorMessage(`Error: ${stderr}`);
            } else {
              vscode.window.showInformationMessage(`File converted to ${finalTexFilePath}`);
            }
          });

          // Pasar el HTML a Pandoc
          if (pandocProcess.stdin) {
            pandocProcess.stdin.write(html);
            pandocProcess.stdin.end();
          } else {
            vscode.window.showErrorMessage('Error: pandocProcess.stdin is null.');
          }
        } catch (error) {
          if (error instanceof Error) {
            vscode.window.showErrorMessage(`Error processing the .docx file: ${error.message}`);
          } else {
            vscode.window.showErrorMessage('An unknown error occurred while processing the .docx file.');
          }
        }
      });
    }
  });

  context.subscriptions.push(exportToWordDisposable);
  context.subscriptions.push(wordToTexDisposable);
}

function extractComments(html: string): { comment: string, position: number }[] {
  const commentRegex = /<sup>(.*?)<\/sup>/g;
  const comments = [];
  let match;
  while ((match = commentRegex.exec(html)) !== null) {
    comments.push({ comment: match[1].trim(), position: match.index });
  }
  return comments;
}

export function deactivate() {}