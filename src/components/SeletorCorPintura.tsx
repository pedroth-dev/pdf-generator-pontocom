import type { CorPinturaPorta, PinturaPorta } from '../context/PortaContext';

/** Grupos por tom: Preto, Branco, Azul (Pintura Automotiva). Cada grupo tem Fosco e Brilhoso. */
const GRUPOS_COR: {
  tone: string;
  swatches: { value: Exclude<CorPinturaPorta, 'Outra' | null>; label: string; hex: string; brilhoso: boolean }[];
}[] = [
  {
    tone: 'Preto',
    swatches: [
      { value: 'Preto Fosco', label: 'Fosco', hex: '#2d2d2d', brilhoso: false },
      { value: 'Preto Brilhoso', label: 'Brilhoso', hex: '#1a1a1a', brilhoso: true },
    ],
  },
  {
    tone: 'Branco',
    swatches: [
      { value: 'Branco Fosco', label: 'Fosco', hex: '#e8e8e8', brilhoso: false },
      { value: 'Branco Brilhoso', label: 'Brilhoso', hex: '#ffffff', brilhoso: true },
    ],
  },
  {
    tone: 'Azul',
    swatches: [
      { value: 'Azul Fosco', label: 'Fosco', hex: '#1e3a5f', brilhoso: false },
      { value: 'Azul Brilhoso', label: 'Brilhoso', hex: '#2563eb', brilhoso: true },
    ],
  },
];

/** Grupos para Primer Dupla Função (cor): Preto, Cinza, Grafite. Cada um com Fosco e Brilhoso. */
const GRUPOS_COR_PRIMER: {
  tone: string;
  swatches: { value: Exclude<CorPinturaPorta, 'Outra' | null>; label: string; hex: string; brilhoso: boolean }[];
}[] = [
  {
    tone: 'Preto',
    swatches: [
      { value: 'Preto Fosco', label: 'Fosco', hex: '#2d2d2d', brilhoso: false },
      { value: 'Preto Brilhoso', label: 'Brilhoso', hex: '#1a1a1a', brilhoso: true },
    ],
  },
  {
    tone: 'Cinza',
    swatches: [
      { value: 'Cinza Fosco', label: 'Fosco', hex: '#6b7280', brilhoso: false },
      { value: 'Cinza Brilhoso', label: 'Brilhoso', hex: '#9ca3af', brilhoso: true },
    ],
  },
  {
    tone: 'Grafite',
    swatches: [
      { value: 'Grafite Fosco', label: 'Fosco', hex: '#374151', brilhoso: false },
      { value: 'Grafite Brilhoso', label: 'Brilhoso', hex: '#4b5563', brilhoso: true },
    ],
  },
];

interface SeletorCorPinturaProps {
  value: CorPinturaPorta;
  corOutra: string;
  onValueChange: (v: CorPinturaPorta) => void;
  onCorOutraChange: (v: string) => void;
  /** Define o conjunto de cores: Automotiva (Preto/Branco/Azul) ou Primer (Preto/Cinza/Grafite). */
  pinturaPorta?: PinturaPorta;
  id?: string;
  error?: string;
}

function Swatch({
  hex,
  brilhoso,
  isSelected,
  onClick,
  label,
  ariaLabel,
}: {
  hex: string;
  brilhoso: boolean;
  isSelected: boolean;
  onClick: () => void;
  label: string;
  ariaLabel: string;
}) {
  const bgStyle = brilhoso
    ? { background: `linear-gradient(135deg, rgba(255,255,255,0.35) 0%, rgba(255,255,255,0) 50%), ${hex}` }
    : { backgroundColor: hex };

  return (
    <button
      type="button"
      onClick={onClick}
      className="flex flex-col items-center gap-1.5 focus:outline-none rounded-[var(--radius)]"
      aria-pressed={isSelected}
      aria-label={ariaLabel}
    >
      <span
        className="w-12 h-12 rounded-lg border-2 transition-all flex-shrink-0"
        style={{
          ...bgStyle,
          borderColor: isSelected ? 'var(--color-accent)' : 'var(--color-border)',
          boxShadow: isSelected ? '0 0 0 2px var(--color-bg)' : undefined,
        }}
      />
      <span className="text-xs text-[var(--color-text-muted)] text-center leading-tight">
        {label}
      </span>
    </button>
  );
}

export function SeletorCorPintura({
  value,
  corOutra,
  onValueChange,
  onCorOutraChange,
  pinturaPorta,
  id,
  error,
}: SeletorCorPinturaProps) {
  const isPrimer = pinturaPorta === 'Primer Dupla Função (cor)';
  const grupos = isPrimer ? GRUPOS_COR_PRIMER : GRUPOS_COR;

  return (
    <div className="flex flex-col gap-4">
      <p className="text-sm font-semibold text-white uppercase tracking-wider">
        Escolha a cor
      </p>
      <div className="grid grid-cols-3 gap-8">
        {grupos.map((grupo) => (
          <div key={grupo.tone} className="flex flex-col items-center gap-2">
            <span className="text-xs font-medium text-[var(--color-text-muted)] uppercase tracking-wider">
              {grupo.tone}
            </span>
            <div className="flex gap-3">
              {grupo.swatches.map((s) => {
                const isSelected = value === s.value;
                return (
                  <Swatch
                    key={s.value}
                    hex={s.hex}
                    brilhoso={s.brilhoso}
                    isSelected={isSelected}
                    onClick={() => onValueChange(s.value)}
                    label={s.label}
                    ariaLabel={`${grupo.tone} ${s.label}`}
                  />
                );
              })}
            </div>
          </div>
        ))}
      </div>
      {/* "Outra" como ação secundária abaixo dos grupos */}
      <div className="flex flex-col gap-2">
        <button
          type="button"
          onClick={() => onValueChange('Outra')}
          className={`inline-flex items-center gap-1.5 self-start py-2 px-3 rounded-[var(--radius)] border text-sm focus:outline-none focus:ring-2 focus:ring-[var(--color-accent)] focus:ring-offset-2 focus:ring-offset-[var(--color-bg)] transition-colors ${
            value === 'Outra'
              ? 'border-[var(--color-accent)] bg-[var(--color-accent-muted)] text-[var(--color-accent)]'
              : 'border-[var(--color-border)] bg-transparent text-[var(--color-text-muted)] hover:bg-[var(--color-surface)] hover:text-white hover:border-[var(--color-text-muted)]'
          }`}
          aria-pressed={value === 'Outra'}
          aria-label="Outra cor"
        >
          <svg className="w-4 h-4 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" aria-hidden>
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
          </svg>
          Outra cor
        </button>
        {value === 'Outra' && (
          <div className="mt-1">
            <label htmlFor={id ? `${id}-outra` : 'cor-pintura-outra'} className="sr-only">
              Especifique a cor
            </label>
            <input
              id={id ? `${id}-outra` : 'cor-pintura-outra'}
              type="text"
              value={corOutra}
              onChange={(e) => onCorOutraChange(e.target.value)}
              placeholder="Ex.: Verde musgo"
              className="w-full max-w-xs py-2.5 px-4 rounded-[var(--radius)] border border-[var(--color-border)] bg-[var(--color-surface)] text-white placeholder-[var(--color-text-muted)] focus:outline-none focus:ring-2 focus:ring-[var(--color-accent)] focus:border-transparent text-sm"
              aria-invalid={!!error}
            />
          </div>
        )}
      </div>
      {error && <p className="text-xs text-amber-400">{error}</p>}
    </div>
  );
}
