import { motion } from 'framer-motion';
import { FileSpreadsheet } from 'lucide-react';

export function Header() {
  return (
    <motion.header
      initial={{ opacity: 0, y: -20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.5 }}
      className="flex items-center gap-3 mb-8"
    >
      <div className="p-2.5 rounded-xl gradient-primary shadow-glow">
        <FileSpreadsheet className="w-7 h-7 text-primary-foreground" />
      </div>
      <div>
        <h1 className="text-2xl font-bold text-foreground tracking-tight">
          Automação Fiscal
        </h1>
        <p className="text-sm text-muted-foreground">
          XML → Excel • NF-e / CT-e
        </p>
      </div>
    </motion.header>
  );
}
