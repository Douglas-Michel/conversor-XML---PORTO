import { useState, useCallback } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { Upload, FileText, CheckCircle2, XCircle, Loader2 } from 'lucide-react';
import { parseNFeXML, NotaFiscal } from '@/lib/xmlParser';
import { DuplicateDialog } from './DuplicateDialog';

interface FileUploadZoneProps {
  onFilesProcessed: (notas: NotaFiscal[]) => void;
  isProcessing: boolean;
  setIsProcessing: (value: boolean) => void;
  existingNotas: NotaFiscal[];
}

export function FileUploadZone({ onFilesProcessed, isProcessing, setIsProcessing, existingNotas }: FileUploadZoneProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [processedCount, setProcessedCount] = useState({ success: 0, failed: 0 });
  const [pendingNotas, setPendingNotas] = useState<NotaFiscal[]>([]);
  const [duplicates, setDuplicates] = useState<NotaFiscal[]>([]);
  const [showDuplicateDialog, setShowDuplicateDialog] = useState(false);

  const processFiles = useCallback(async (files: FileList | File[]) => {
    setIsProcessing(true);
    setProcessedCount({ success: 0, failed: 0 });

    const xmlFiles = Array.from(files).filter(
      file => file.name.toLowerCase().endsWith('.xml')
    );

    if (xmlFiles.length === 0) {
      setIsProcessing(false);
      return;
    }

    const notas: NotaFiscal[] = [];
    let success = 0;
    let failed = 0;

    for (const file of xmlFiles) {
      try {
        const content = await file.text();
        const nota = parseNFeXML(content, file.name);
        if (nota) {
          notas.push(nota);
          success++;
        } else {
          failed++;
        }
      } catch (error) {
        console.error(`Error processing ${file.name}:`, error);
        failed++;
      }
      setProcessedCount({ success, failed });
    }

    // Check for duplicates based on chaveAcesso
    const existingKeys = new Set(existingNotas.map(n => n.chaveAcesso));
    const duplicateNotas = notas.filter(n => existingKeys.has(n.chaveAcesso));
    const uniqueNotas = notas.filter(n => !existingKeys.has(n.chaveAcesso));

    if (duplicateNotas.length > 0) {
      setPendingNotas(uniqueNotas);
      setDuplicates(duplicateNotas);
      setShowDuplicateDialog(true);
      setIsProcessing(false);
    } else {
      onFilesProcessed(notas);
      setIsProcessing(false);
    }
  }, [onFilesProcessed, setIsProcessing, existingNotas]);

  const handleConfirmDuplicates = () => {
    // Import all including duplicates
    onFilesProcessed([...pendingNotas, ...duplicates]);
    setShowDuplicateDialog(false);
    setPendingNotas([]);
    setDuplicates([]);
  };

  const handleCancelDuplicates = () => {
    // Import only non-duplicates
    if (pendingNotas.length > 0) {
      onFilesProcessed(pendingNotas);
    }
    setShowDuplicateDialog(false);
    setPendingNotas([]);
    setDuplicates([]);
  };

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      processFiles(files);
    }
  }, [processFiles]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      processFiles(files);
    }
    e.target.value = '';
  }, [processFiles]);

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.4, ease: 'easeOut' }}
    >
      <label
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        className={`
          relative flex flex-col items-center justify-center 
          w-full min-h-[280px] p-8
          border-2 border-dashed rounded-xl cursor-pointer
          transition-all duration-300 ease-out
          ${isDragging 
            ? 'border-primary bg-primary/6 scale-[1.02] shadow-glow' 
            : 'border-border bg-card hover:border-primary/50 hover:bg-muted/50'
          }
          ${isProcessing ? 'pointer-events-none' : ''}
        `}
      >
        <input
          type="file"
          accept=".xml"
          multiple
          onChange={handleFileSelect}
          className="sr-only"
          disabled={isProcessing}
        />

        <AnimatePresence mode="wait">
          {isProcessing ? (
            <motion.div
              key="processing"
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              className="flex flex-col items-center gap-4"
            >
              <Loader2 className="w-12 h-12 text-primary animate-spin" />
              <div className="text-center">
                <p className="text-lg font-medium text-foreground">
                  Processando arquivos...
                </p>
                <div className="flex items-center justify-center gap-4 mt-2 text-sm text-muted-foreground">
                  <span className="flex items-center gap-1">
                    <CheckCircle2 className="w-4 h-4 text-success" />
                    {processedCount.success} sucesso
                  </span>
                  {processedCount.failed > 0 && (
                    <span className="flex items-center gap-1">
                      <XCircle className="w-4 h-4 text-destructive" />
                      {processedCount.failed} falha
                    </span>
                  )}
                </div>
              </div>
            </motion.div>
          ) : (
            <motion.div
              key="idle"
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              className="flex flex-col items-center gap-4 text-center"
            >
              <div className={`
                p-4 rounded-full transition-colors duration-300
                ${isDragging ? 'bg-primary/10 shadow-glow' : 'bg-muted/20'}
              `}>
                {isDragging ? (
                  <FileText className="w-10 h-10 text-primary" />
                ) : (
                  <Upload className="w-10 h-10 text-muted-foreground" />
                )}
              </div>
              
              <div>
                <p className="text-lg font-medium text-foreground">
                  {isDragging ? 'Solte os arquivos aqui' : 'Arraste arquivos XML aqui'}
                </p>
                <p className="mt-1 text-sm text-muted-foreground">
                  ou clique para selecionar
                </p>
              </div>

              <div className="flex items-center gap-2 px-4 py-2 mt-2 rounded-lg bg-muted/20 border border-border">
                <FileText className="w-4 h-4 text-muted-foreground" />
                <span className="text-xs font-medium text-muted-foreground">
                  NF-e • CT-e • Múltiplos arquivos
                </span>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </label>

      <DuplicateDialog
        open={showDuplicateDialog}
        duplicates={duplicates}
        onConfirm={handleConfirmDuplicates}
        onCancel={handleCancelDuplicates}
      />
    </motion.div>
  );
}
