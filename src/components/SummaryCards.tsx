import { motion } from 'framer-motion';
import { FileText, ArrowDownCircle, ArrowUpCircle, Calculator } from 'lucide-react';
import { NotaFiscal, formatCurrency } from '@/lib/xmlParser';

interface SummaryCardsProps {
  data: NotaFiscal[];
}

export function SummaryCards({ data }: SummaryCardsProps) {
  if (data.length === 0) return null;

  const entradas = data.filter(n => n.tipoOperacao === 'Entrada');
  const saidas = data.filter(n => n.tipoOperacao === 'Saída');
  
  const totalEntradas = entradas.reduce((sum, n) => sum + n.valorTotal, 0);
  const totalSaidas = saidas.reduce((sum, n) => sum + n.valorTotal, 0);
  const totalICMS = data.reduce((sum, n) => sum + n.valorICMS, 0);

  const cards = [
    {
      title: 'Total Documentos',
      value: data.length.toString(),
      subtitle: `${data.filter(n => n.tipo === 'NF-e').length} NF-e • ${data.filter(n => n.tipo === 'CT-e').length} CT-e`,
      icon: FileText,
      color: 'primary',
    },
    {
      title: 'Entradas',
      value: formatCurrency(totalEntradas),
      subtitle: `${entradas.length} documento${entradas.length !== 1 ? 's' : ''}`,
      icon: ArrowDownCircle,
      color: 'success',
    },
    {
      title: 'Saídas',
      value: formatCurrency(totalSaidas),
      subtitle: `${saidas.length} documento${saidas.length !== 1 ? 's' : ''}`,
      icon: ArrowUpCircle,
      color: 'accent',
    },
    {
      title: 'Total ICMS',
      value: formatCurrency(totalICMS),
      subtitle: 'Soma de todos os documentos',
      icon: Calculator,
      color: 'muted',
    },
  ];

  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
      {cards.map((card, index) => (
        <motion.div
          key={card.title}
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.4, delay: index * 0.1 }}
          className="p-5 rounded-xl bg-card border border-border shadow-soft hover:shadow-elevated transition-shadow"
        >
          <div className="flex items-start justify-between">
            <div className="flex-1">
              <p className="text-sm font-medium text-muted-foreground">{card.title}</p>
              <p className="mt-2 text-2xl font-semibold tracking-tight text-foreground">
                {card.value}
              </p>
              <p className="mt-1 text-xs text-muted-foreground">{card.subtitle}</p>
            </div>
            <div className={`
              p-2.5 rounded-lg
              ${card.color === 'primary' ? 'bg-primary/10 text-primary' : ''}
              ${card.color === 'success' ? 'bg-success/10 text-success' : ''}
              ${card.color === 'accent' ? 'bg-accent/10 text-accent' : ''}
              ${card.color === 'muted' ? 'bg-muted text-muted-foreground' : ''}
            `}>
              <card.icon className="w-5 h-5" />
            </div>
          </div>
        </motion.div>
      ))}
    </div>
  );
}
