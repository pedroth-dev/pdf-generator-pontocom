import {
  createContext,
  useContext,
  useReducer,
  useCallback,
  type ReactNode,
} from 'react';

export type ModeloPorta =
  | 'Boiserie'
  | 'Boiserie e Ferro Forjado'
  | 'Ferro Forjado'
  | 'Bandeja'
  | 'Aço Corten'
  | 'Lisa'
  | null;

export type ModoPuxador =
  | 'Puxador Cava'
  | 'Puxador Sobrepor no Metalon 40x20'
  | null;

export type SistemaAbertura =
  | 'Pivotante'
  | 'De Correr'
  | 'Com Dobradiça'
  | null;

export type EstiloFolha = 'Folha Única' | 'Folha Dupla' | null;

export type AcondicionamentoPorta = 'Almofadada' | 'Bandeja Simples' | null;

export type PinturaPorta =
  | 'Pintura Automotiva (cor)'
  | 'Pintura Corten'
  | 'Aço Corten'
  | 'Primer Dupla Função (cor)'
  | null;

export type CorPinturaPorta =
  | 'Preto Fosco'
  | 'Preto Brilhoso'
  | 'Branco Fosco'
  | 'Branco Brilhoso'
  | 'Azul Fosco'
  | 'Azul Brilhoso'
  | 'Cinza Fosco'
  | 'Cinza Brilhoso'
  | 'Grafite Fosco'
  | 'Grafite Brilhoso'
  | 'Outra'
  | null;

export type ModoEntrega = 'Instalada no Local' | 'Retirada na Empresa' | null;

/** Espessura da chapa por modelo (valor fixo, não editável) */
export const ESPESSURA_CHAPA_POR_MODELO: Record<NonNullable<ModeloPorta>, string> = {
  Boiserie: '#14',
  'Boiserie e Ferro Forjado': '#14',
  'Ferro Forjado': '#14',
  Lisa: '#14',
  'Aço Corten': 'aço cor 420',
  Bandeja: '#18',
};

export interface PortaFormData {
  modeloPorta: NonNullable<ModeloPorta>;
  modoPuxador: NonNullable<ModoPuxador>;
  sistemaAbertura: NonNullable<SistemaAbertura>;
  estiloFolha: NonNullable<EstiloFolha>;
  acondicionamentoEfetivo: string; // "almofadada" | "almofadada" | "bandeja simples" | "" (vazio para Ferro/Aço Corten)
  espessuraChapa: string;
  bandeirola: boolean;
  alizar: boolean;
  pinturaPorta: NonNullable<PinturaPorta>;
  corPintura: string; // cor para substituir (cor) ou nome da pintura sem cor (Corten, Aço Corten)
  valorM2: string;
  medidasPorta: string; // "2,00m x 3,00m"
  alturaPorta: number;
  larguraPorta: number;
  medidasPortaGeral: string;
  m2: number;
  alturaBandeirola?: number;
  larguraBandeirola?: number;
  medidasBandeirola?: string;
  medidaAlizar?: string;
  custoDeslocamento: string;
  descricaoAdicional?: string;
  modoEntrega: NonNullable<ModoEntrega>;
}

interface PortaState {
  modeloPorta: ModeloPorta;
  modoPuxador: ModoPuxador;
  sistemaAbertura: SistemaAbertura;
  estiloFolha: EstiloFolha;
  acondicionamento: AcondicionamentoPorta;
  pinturaPorta: PinturaPorta;
  corPintura: CorPinturaPorta;
  corPinturaOutra: string;
  modoEntrega: ModoEntrega;
  bandeirola: boolean;
  alizar: boolean;
  valorM2: string;
  alturaPorta: string;
  larguraPorta: string;
  alturaBandeirola: string;
  larguraBandeirola: string;
  medidaAlizar: string;
  custoDeslocamento: string;
  descricaoAdicional: string;
  formTouched: boolean;
}

type PortaAction =
  | { type: 'SET_MODELO'; payload: ModeloPorta }
  | { type: 'SET_MODO_PUXADOR'; payload: ModoPuxador }
  | { type: 'SET_SISTEMA_ABERTURA'; payload: SistemaAbertura }
  | { type: 'SET_ESTILO_FOLHA'; payload: EstiloFolha }
  | { type: 'SET_ACONDICIONAMENTO'; payload: AcondicionamentoPorta }
  | { type: 'SET_PINTURA'; payload: PinturaPorta }
  | { type: 'SET_COR_PINTURA'; payload: CorPinturaPorta }
  | { type: 'SET_COR_PINTURA_OUTRA'; payload: string }
  | { type: 'SET_MODO_ENTREGA'; payload: ModoEntrega }
  | { type: 'SET_BANDEIROLA'; payload: boolean }
  | { type: 'SET_ALIZAR'; payload: boolean }
  | { type: 'SET_VALOR_M2'; payload: string }
  | { type: 'SET_ALTURA_PORTA'; payload: string }
  | { type: 'SET_LARGURA_PORTA'; payload: string }
  | { type: 'SET_ALTURA_BANDEIROLA'; payload: string }
  | { type: 'SET_LARGURA_BANDEIROLA'; payload: string }
  | { type: 'SET_MEDIDA_ALIZAR'; payload: string }
  | { type: 'SET_CUSTO_DESLOCAMENTO'; payload: string }
  | { type: 'SET_DESCRICAO_ADICIONAL'; payload: string }
  | { type: 'SET_FORM_TOUCHED'; payload: boolean }
  | { type: 'RESET' };

const initialState: PortaState = {
  modeloPorta: null,
  modoPuxador: null,
  sistemaAbertura: null,
  estiloFolha: null,
  acondicionamento: null,
  pinturaPorta: null,
  corPintura: null,
  corPinturaOutra: '',
  modoEntrega: null,
  bandeirola: false,
  alizar: false,
  valorM2: '',
  alturaPorta: '',
  larguraPorta: '',
  alturaBandeirola: '',
  larguraBandeirola: '',
  medidaAlizar: '',
  custoDeslocamento: '',
  descricaoAdicional: '',
  formTouched: false,
};

function parseMetro(s: string): number {
  const v = parseFloat((s || '').trim().replace(',', '.'));
  return Number.isFinite(v) && v >= 0 ? v : 0;
}

function formatMedidas(altura: number, largura: number): string {
  if (altura <= 0 || largura <= 0) return '';
  return `${altura.toFixed(2).replace('.', ',')}m x ${largura.toFixed(2).replace('.', ',')}m`;
}

function portaReducer(state: PortaState, action: PortaAction): PortaState {
  switch (action.type) {
    case 'SET_MODELO': {
      const modelo = action.payload;
      const next = { ...state, modeloPorta: modelo };
      if (modelo && modelo !== 'Lisa' && modelo !== 'Aço Corten') {
        next.modoPuxador = 'Puxador Sobrepor no Metalon 40x20';
      }
      if (modelo === 'Bandeja' && state.acondicionamento === null) {
        next.acondicionamento = 'Almofadada';
      }
      if (modelo !== 'Bandeja') {
        next.acondicionamento = null;
      }
      return next;
    }
    case 'SET_MODO_PUXADOR':
      return { ...state, modoPuxador: action.payload };
    case 'SET_SISTEMA_ABERTURA':
      return { ...state, sistemaAbertura: action.payload };
    case 'SET_ESTILO_FOLHA':
      return { ...state, estiloFolha: action.payload };
    case 'SET_ACONDICIONAMENTO':
      return { ...state, acondicionamento: action.payload };
    case 'SET_PINTURA':
      return { ...state, pinturaPorta: action.payload };
    case 'SET_COR_PINTURA':
      return { ...state, corPintura: action.payload };
    case 'SET_COR_PINTURA_OUTRA':
      return { ...state, corPinturaOutra: action.payload };
    case 'SET_MODO_ENTREGA':
      return { ...state, modoEntrega: action.payload };
    case 'SET_BANDEIROLA':
      return { ...state, bandeirola: action.payload };
    case 'SET_ALIZAR':
      return { ...state, alizar: action.payload };
    case 'SET_VALOR_M2':
      return { ...state, valorM2: action.payload };
    case 'SET_ALTURA_PORTA':
      return { ...state, alturaPorta: action.payload };
    case 'SET_LARGURA_PORTA':
      return { ...state, larguraPorta: action.payload };
    case 'SET_ALTURA_BANDEIROLA':
      return { ...state, alturaBandeirola: action.payload };
    case 'SET_LARGURA_BANDEIROLA':
      return { ...state, larguraBandeirola: action.payload };
    case 'SET_MEDIDA_ALIZAR':
      return { ...state, medidaAlizar: action.payload };
    case 'SET_CUSTO_DESLOCAMENTO':
      return { ...state, custoDeslocamento: action.payload };
    case 'SET_DESCRICAO_ADICIONAL':
      return { ...state, descricaoAdicional: action.payload };
    case 'SET_FORM_TOUCHED':
      return { ...state, formTouched: action.payload };
    case 'RESET':
      return initialState;
    default:
      return state;
  }
}

interface PortaContextValue extends PortaState {
  setModeloPorta: (v: ModeloPorta) => void;
  setModoPuxador: (v: ModoPuxador) => void;
  setSistemaAbertura: (v: SistemaAbertura) => void;
  setEstiloFolha: (v: EstiloFolha) => void;
  setAcondicionamento: (v: AcondicionamentoPorta) => void;
  setPinturaPorta: (v: PinturaPorta) => void;
  setCorPintura: (v: CorPinturaPorta) => void;
  setCorPinturaOutra: (v: string) => void;
  setModoEntrega: (v: ModoEntrega) => void;
  setBandeirola: (v: boolean) => void;
  setAlizar: (v: boolean) => void;
  setValorM2: (v: string) => void;
  setAlturaPorta: (v: string) => void;
  setLarguraPorta: (v: string) => void;
  setAlturaBandeirola: (v: string) => void;
  setLarguraBandeirola: (v: string) => void;
  setMedidaAlizar: (v: string) => void;
  setCustoDeslocamento: (v: string) => void;
  setDescricaoAdicional: (v: string) => void;
  setFormTouched: (v: boolean) => void;
  reset: () => void;
  modoPuxadorLocked: boolean;
  espessuraChapa: string;
  acondicionamentoVisivel: boolean;
  acondicionamentoEfetivo: string;
  medidasPortaGeral: string;
  m2: number;
  buildFormData: () => PortaFormData | null;
}

const PortaContext = createContext<PortaContextValue | null>(null);

export function PortaProvider({ children }: { children: ReactNode }) {
  const [state, dispatch] = useReducer(portaReducer, initialState);

  const setModeloPorta = useCallback((p: ModeloPorta) => dispatch({ type: 'SET_MODELO', payload: p }), []);
  const setModoPuxador = useCallback((p: ModoPuxador) => dispatch({ type: 'SET_MODO_PUXADOR', payload: p }), []);
  const setSistemaAbertura = useCallback((p: SistemaAbertura) => dispatch({ type: 'SET_SISTEMA_ABERTURA', payload: p }), []);
  const setEstiloFolha = useCallback((p: EstiloFolha) => dispatch({ type: 'SET_ESTILO_FOLHA', payload: p }), []);
  const setAcondicionamento = useCallback((p: AcondicionamentoPorta) => dispatch({ type: 'SET_ACONDICIONAMENTO', payload: p }), []);
  const setPinturaPorta = useCallback((p: PinturaPorta) => dispatch({ type: 'SET_PINTURA', payload: p }), []);
  const setCorPintura = useCallback((p: CorPinturaPorta) => dispatch({ type: 'SET_COR_PINTURA', payload: p }), []);
  const setCorPinturaOutra = useCallback((p: string) => dispatch({ type: 'SET_COR_PINTURA_OUTRA', payload: p }), []);
  const setModoEntrega = useCallback((p: ModoEntrega) => dispatch({ type: 'SET_MODO_ENTREGA', payload: p }), []);
  const setBandeirola = useCallback((p: boolean) => dispatch({ type: 'SET_BANDEIROLA', payload: p }), []);
  const setAlizar = useCallback((p: boolean) => dispatch({ type: 'SET_ALIZAR', payload: p }), []);
  const setValorM2 = useCallback((p: string) => dispatch({ type: 'SET_VALOR_M2', payload: p }), []);
  const setAlturaPorta = useCallback((p: string) => dispatch({ type: 'SET_ALTURA_PORTA', payload: p }), []);
  const setLarguraPorta = useCallback((p: string) => dispatch({ type: 'SET_LARGURA_PORTA', payload: p }), []);
  const setAlturaBandeirola = useCallback((p: string) => dispatch({ type: 'SET_ALTURA_BANDEIROLA', payload: p }), []);
  const setLarguraBandeirola = useCallback((p: string) => dispatch({ type: 'SET_LARGURA_BANDEIROLA', payload: p }), []);
  const setMedidaAlizar = useCallback((p: string) => dispatch({ type: 'SET_MEDIDA_ALIZAR', payload: p }), []);
  const setCustoDeslocamento = useCallback((p: string) => dispatch({ type: 'SET_CUSTO_DESLOCAMENTO', payload: p }), []);
  const setDescricaoAdicional = useCallback((p: string) => dispatch({ type: 'SET_DESCRICAO_ADICIONAL', payload: p }), []);
  const setFormTouched = useCallback((p: boolean) => dispatch({ type: 'SET_FORM_TOUCHED', payload: p }), []);
  const reset = useCallback(() => dispatch({ type: 'RESET' }), []);

  const modoPuxadorLocked =
    state.modeloPorta !== null && state.modeloPorta !== 'Lisa' && state.modeloPorta !== 'Aço Corten';
  const espessuraChapa = state.modeloPorta ? ESPESSURA_CHAPA_POR_MODELO[state.modeloPorta] : '';
  const acondicionamentoVisivel = state.modeloPorta === 'Bandeja';
  const acondicionamentoEfetivo = (() => {
    if (!state.modeloPorta) return '';
    if (state.modeloPorta === 'Bandeja') return (state.acondicionamento || 'Almofadada').toLowerCase().replace(' ', ' ');
    if (['Boiserie', 'Boiserie e Ferro Forjado', 'Lisa'].includes(state.modeloPorta)) return 'almofadada';
    return ''; // Ferro Forjado, Aço Corten: sem placeholder na descrição
  })();

  const altPorta = parseMetro(state.alturaPorta);
  const largPorta = parseMetro(state.larguraPorta);
  const altBandeira = state.bandeirola ? parseMetro(state.alturaBandeirola) : 0;
  const largBandeira = state.bandeirola ? parseMetro(state.larguraBandeirola) : 0;
  const alturaGeral = altPorta + altBandeira;
  const larguraGeral = largPorta + largBandeira;
  const medidasPortaGeral = alturaGeral > 0 && larguraGeral > 0 ? formatMedidas(alturaGeral, larguraGeral) : '';
  const m2 = Math.round(alturaGeral * larguraGeral * 100) / 100;

  const buildFormData = useCallback((): PortaFormData | null => {
    if (
      !state.modeloPorta ||
      !state.modoPuxador ||
      !state.sistemaAbertura ||
      !state.estiloFolha ||
      !state.pinturaPorta ||
      !state.modoEntrega
    ) {
      return null;
    }
    const medidasPortaStr = altPorta > 0 && largPorta > 0 ? formatMedidas(altPorta, largPorta) : '';
    const precisaCor = state.pinturaPorta === 'Pintura Automotiva (cor)' || state.pinturaPorta === 'Primer Dupla Função (cor)';
    const corTexto =
      precisaCor && state.corPintura
        ? state.corPintura === 'Outra'
          ? state.corPinturaOutra.trim() || ''
          : state.corPintura
        : state.pinturaPorta === 'Pintura Corten'
          ? 'pintura corten'
          : state.pinturaPorta === 'Aço Corten'
            ? 'aço corten'
            : '';
    const pinturaComCor =
      state.pinturaPorta === 'Pintura Automotiva (cor)' || state.pinturaPorta === 'Primer Dupla Função (cor)'
        ? (state.pinturaPorta === 'Pintura Automotiva (cor)' ? 'pintura automotiva ' : 'primer dupla função ') + (corTexto.toLowerCase())
        : corTexto;

    const base: PortaFormData = {
      modeloPorta: state.modeloPorta,
      modoPuxador: state.modoPuxador,
      sistemaAbertura: state.sistemaAbertura,
      estiloFolha: state.estiloFolha,
      acondicionamentoEfetivo,
      espessuraChapa,
      bandeirola: state.bandeirola,
      alizar: state.alizar,
      pinturaPorta: state.pinturaPorta,
      corPintura: pinturaComCor,
      valorM2: state.valorM2,
      medidasPorta: medidasPortaStr,
      alturaPorta: altPorta,
      larguraPorta: largPorta,
      medidasPortaGeral,
      m2,
      custoDeslocamento: state.custoDeslocamento,
      modoEntrega: state.modoEntrega,
    };
    if (state.descricaoAdicional.trim()) base.descricaoAdicional = state.descricaoAdicional.trim();
    if (state.bandeirola) {
      base.alturaBandeirola = altBandeira;
      base.larguraBandeirola = largBandeira;
      base.medidasBandeirola = altBandeira > 0 && largBandeira > 0 ? formatMedidas(altBandeira, largBandeira) : '';
    }
    if (state.alizar && state.medidaAlizar.trim()) {
      const v = parseMetro(state.medidaAlizar);
      base.medidaAlizar = v > 0 ? `${v.toFixed(2).replace('.', ',')}m` : state.medidaAlizar.trim();
    }
    return base;
  }, [
    state.modeloPorta,
    state.modoPuxador,
    state.sistemaAbertura,
    state.estiloFolha,
    state.pinturaPorta,
    state.corPintura,
    state.corPinturaOutra,
    state.modoEntrega,
    state.bandeirola,
    state.alizar,
    state.valorM2,
    state.alturaPorta,
    state.larguraPorta,
    state.alturaBandeirola,
    state.larguraBandeirola,
    state.medidaAlizar,
    state.custoDeslocamento,
    state.descricaoAdicional,
    acondicionamentoEfetivo,
    espessuraChapa,
    medidasPortaGeral,
    m2,
    altPorta,
    largPorta,
    altBandeira,
    largBandeira,
  ]);

  const value: PortaContextValue = {
    ...state,
    setModeloPorta,
    setModoPuxador,
    setSistemaAbertura,
    setEstiloFolha,
    setAcondicionamento,
    setPinturaPorta,
    setCorPintura,
    setCorPinturaOutra,
    setModoEntrega,
    setBandeirola,
    setAlizar,
    setValorM2,
    setAlturaPorta,
    setLarguraPorta,
    setAlturaBandeirola,
    setLarguraBandeirola,
    setMedidaAlizar,
    setCustoDeslocamento,
    setDescricaoAdicional,
    setFormTouched,
    reset,
    modoPuxadorLocked,
    espessuraChapa,
    acondicionamentoVisivel,
    acondicionamentoEfetivo,
    medidasPortaGeral,
    m2,
    buildFormData,
  };

  return (
    <PortaContext.Provider value={value}>
      {children}
    </PortaContext.Provider>
  );
}

export function usePorta() {
  const ctx = useContext(PortaContext);
  if (!ctx) throw new Error('usePorta must be used within PortaProvider');
  return ctx;
}
