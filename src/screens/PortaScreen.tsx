import { useMemo } from 'react';
import { usePorta } from '../context/PortaContext';
import type { PortaFormData } from '../context/PortaContext';
import { StepButtons } from '../components/StepButtons';
import { CurrencyInput } from '../components/CurrencyInput';
import { MedidasAlturaLarguraInput, parseAlturaLargura } from '../components/MedidasAlturaLarguraInput';
import { SeletorCorPintura } from '../components/SeletorCorPintura';

interface PortaScreenProps {
  onBack: () => void;
  onConfirm: (data: PortaFormData) => void;
}

const MODELOS = [
  { value: 'Boiserie', label: 'Boiserie' },
  { value: 'Boiserie e Ferro Forjado', label: 'Boiserie e Ferro Forjado' },
  { value: 'Ferro Forjado', label: 'Ferro Forjado' },
  { value: 'Bandeja', label: 'Bandeja' },
  { value: 'Aço Corten', label: 'Aço Corten' },
  { value: 'Lisa', label: 'Lisa' },
] as const;

const MODO_PUXADOR_OPCOES = [
  { value: 'Puxador Cava', label: 'Puxador Cava' },
  { value: 'Puxador Sobrepor no Metalon 40x20', label: 'Puxador Sobrepor no Metalon 40x20' },
] as const;

const SISTEMA_ABERTURA_OPCOES = [
  { value: 'Pivotante', label: 'Pivotante' },
  { value: 'De Correr', label: 'De Correr' },
  { value: 'Com Dobradiça', label: 'Com Dobradiça' },
] as const;

const ESTILO_FOLHA_OPCOES = [
  { value: 'Folha Única', label: 'Folha Única' },
  { value: 'Folha Dupla', label: 'Folha Dupla' },
] as const;

const ACONDICIONAMENTO_OPCOES = [
  { value: 'Almofadada', label: 'Almofadada' },
  { value: 'Bandeja Simples', label: 'Bandeja Simples' },
] as const;

const PINTURA_OPCOES = [
  { value: 'Pintura Automotiva (cor)', label: 'Pintura Automotiva (cor)' },
  { value: 'Pintura Corten', label: 'Pintura Corten' },
  { value: 'Aço Corten', label: 'Aço Corten' },
  { value: 'Primer Dupla Função (cor)', label: 'Primer Dupla Função (cor)' },
] as const;

const MODO_ENTREGA_OPCOES = [
  { value: 'Instalada no Local', label: 'Instalada no Local' },
  { value: 'Retirada na Empresa', label: 'Retirada na Empresa' },
] as const;

function parseMedidaSimples(s: string): number {
  const v = parseFloat((s || '').trim().replace(',', '.'));
  return Number.isFinite(v) && v >= 0 ? v : 0;
}

export function PortaScreen({ onBack, onConfirm }: PortaScreenProps) {
  const {
    modeloPorta,
    modoPuxador,
    modoPuxadorLocked,
    sistemaAbertura,
    estiloFolha,
    acondicionamento,
    acondicionamentoVisivel,
    pinturaPorta,
    corPintura,
    corPinturaOutra,
    modoEntrega,
    bandeirola,
    alizar,
    valorM2,
    alturaPorta,
    larguraPorta,
    alturaBandeirola,
    larguraBandeirola,
    medidaAlizar,
    custoDeslocamento,
    descricaoAdicional,
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
    buildFormData,
    formTouched,
    setFormTouched,
  } = usePorta();

  const precisaCor = pinturaPorta === 'Pintura Automotiva (cor)' || pinturaPorta === 'Primer Dupla Função (cor)';
  const corOk = !precisaCor || (corPintura !== null && (corPintura !== 'Outra' || corPinturaOutra.trim().length > 0));

  /** Opções de pintura com rótulo resolvido quando uma cor está selecionada (ex.: "Pintura Automotiva — Preto Fosco"). */
  const pinturaOpcoesComResolucao = useMemo(() => {
    return PINTURA_OPCOES.map((opt) => {
      const isComCor = opt.value === 'Pintura Automotiva (cor)' || opt.value === 'Primer Dupla Função (cor)';
      if (!isComCor || opt.value !== pinturaPorta) return { ...opt };
      const nomeCor = corPintura === 'Outra' ? corPinturaOutra.trim() : (corPintura || '');
      if (!nomeCor) return { ...opt };
      const base = opt.value === 'Pintura Automotiva (cor)' ? 'Pintura Automotiva' : 'Primer Dupla Função';
      return { ...opt, label: `${base} — ${nomeCor}` };
    });
  }, [pinturaPorta, corPintura, corPinturaOutra]);

  const { altura: altP, largura: largP } = parseAlturaLargura(alturaPorta, larguraPorta);
  const okMedidasPorta = altP > 0 && largP > 0;
  const { altura: altB, largura: largB } = parseAlturaLargura(alturaBandeirola, larguraBandeirola);
  const okMedidasBandeirola = !bandeirola || (altB > 0 && largB > 0);
  const medidaAlizarNum = parseMedidaSimples(medidaAlizar);
  const okMedidaAlizar = !alizar || medidaAlizarNum > 0;
  const okValorM2 = valorM2.length > 0 && parseInt(valorM2, 10) > 0;
  const okCustoDeslocamento = custoDeslocamento.length > 0 && parseInt(custoDeslocamento, 10) >= 0;

  const canSubmit =
    modeloPorta !== null &&
    modoPuxador !== null &&
    sistemaAbertura !== null &&
    estiloFolha !== null &&
    (!acondicionamentoVisivel || acondicionamento !== null) &&
    pinturaPorta !== null &&
    corOk &&
    modoEntrega !== null &&
    okMedidasPorta &&
    okMedidasBandeirola &&
    okMedidaAlizar &&
    okValorM2 &&
    okCustoDeslocamento;

  const handleSubmit = () => {
    setFormTouched(true);
    if (!canSubmit) return;
    const data = buildFormData();
    if (data) onConfirm(data);
  };

  return (
    <div className="min-h-full flex flex-col px-6 py-8 max-w-5xl mx-auto w-full">
      <button
        type="button"
        onClick={onBack}
        className="self-start text-sm font-medium text-[var(--color-text-muted)] hover:text-[var(--color-accent)] transition-colors mb-8"
      >
        ← Voltar
      </button>

      <h2 className="text-xl font-semibold text-white mb-1">Porta</h2>
      <p className="text-sm text-[var(--color-text-muted)] mb-2">Dados do orçamento</p>
      <p className="text-sm text-[var(--color-text-muted)] mb-8">
        Preencha as opções e os valores. A ordem das medidas é altura e depois largura.
      </p>

      {/* Seção 1 — Opções */}
      <div className="mb-8">
        <p className="text-sm font-medium text-white mb-3">1. Modelo da Porta</p>
        <StepButtons
          options={[...MODELOS]}
          value={modeloPorta}
          onChange={(v) => setModeloPorta(v as typeof modeloPorta)}
        />
      </div>

      <div className="mb-8">
        <p className="text-sm font-medium text-white mb-3">2. Modo de Puxador</p>
        {modoPuxadorLocked ? (
          <div className="py-3 px-0 text-[var(--color-text-muted)] w-fit" aria-hidden>
            Puxador Sobrepor no Metalon 40x20
          </div>
        ) : (
          <StepButtons
            options={[...MODO_PUXADOR_OPCOES]}
            value={modeloPorta === 'Lisa' || modeloPorta === 'Aço Corten' ? modoPuxador : null}
            onChange={(v) => setModoPuxador(v as typeof modoPuxador)}
            disabled={modeloPorta !== 'Lisa' && modeloPorta !== 'Aço Corten'}
          />
        )}
      </div>

      <div className="mb-8">
        <p className="text-sm font-medium text-white mb-3">3. Sistema de Abertura</p>
        <StepButtons
          options={[...SISTEMA_ABERTURA_OPCOES]}
          value={sistemaAbertura}
          onChange={(v) => setSistemaAbertura(v as typeof sistemaAbertura)}
        />
      </div>

      <div className="mb-8">
        <p className="text-sm font-medium text-white mb-3">4. Estilo de Folha</p>
        <StepButtons
          options={[...ESTILO_FOLHA_OPCOES]}
          value={estiloFolha}
          onChange={(v) => setEstiloFolha(v as typeof estiloFolha)}
        />
      </div>

      {acondicionamentoVisivel && (
        <div className="mb-8">
          <p className="text-sm font-medium text-white mb-3">5. Acondicionamento</p>
          <StepButtons
            options={[...ACONDICIONAMENTO_OPCOES]}
            value={acondicionamento}
            onChange={(v) => setAcondicionamento(v as typeof acondicionamento)}
          />
        </div>
      )}

      {/* Seção 2 — Itens complementares */}
      <div className="mb-8">
        <p className="text-sm font-medium text-white mb-3">Itens complementares</p>
        <p className="text-xs text-[var(--color-text-muted)] mb-3">
          Selecione nenhum, um ou ambos.
        </p>
        <div className="flex flex-wrap gap-3">
          <button
            type="button"
            onClick={() => setBandeirola(!bandeirola)}
            aria-pressed={bandeirola}
            className={`py-3 px-5 rounded-[var(--radius)] border font-medium transition-all duration-200 focus:outline-none focus:ring-2 focus:ring-[var(--color-accent)] focus:ring-offset-2 focus:ring-offset-[var(--color-bg)] ${
              bandeirola
                ? 'border-[var(--color-accent)] bg-[var(--color-accent-muted)] text-[var(--color-accent)]'
                : 'border-[var(--color-border)] bg-[var(--color-surface)] text-white hover:border-[var(--color-text-muted)]'
            }`}
          >
            Bandeirola
          </button>
          <button
            type="button"
            onClick={() => setAlizar(!alizar)}
            aria-pressed={alizar}
            className={`py-3 px-5 rounded-[var(--radius)] border font-medium transition-all duration-200 focus:outline-none focus:ring-2 focus:ring-[var(--color-accent)] focus:ring-offset-2 focus:ring-offset-[var(--color-bg)] ${
              alizar
                ? 'border-[var(--color-accent)] bg-[var(--color-accent-muted)] text-[var(--color-accent)]'
                : 'border-[var(--color-border)] bg-[var(--color-surface)] text-white hover:border-[var(--color-text-muted)]'
            }`}
          >
            Alizar
          </button>
        </div>
      </div>

      {/* Seção 3 — Pintura */}
      <div className="mb-8">
        <p className="text-sm font-medium text-white mb-3">Pintura da Porta</p>
        <StepButtons
          options={pinturaOpcoesComResolucao}
          value={pinturaPorta}
          onChange={(v) => setPinturaPorta(v as typeof pinturaPorta)}
        />
        {precisaCor && (
          <div className="mt-4 pl-0">
            <SeletorCorPintura
              value={corPintura}
              corOutra={corPinturaOutra}
              onValueChange={setCorPintura}
              onCorOutraChange={setCorPinturaOutra}
              pinturaPorta={pinturaPorta ?? undefined}
              id="porta-cor"
              error={formTouched && precisaCor && !corOk ? 'Selecione ou informe a cor.' : undefined}
            />
          </div>
        )}
      </div>

      {/* Seção 4 — Modo de Entrega */}
      <div className="mb-8">
        <p className="text-sm font-medium text-white mb-3">Modo de Entrega</p>
        <StepButtons
          options={[...MODO_ENTREGA_OPCOES]}
          value={modoEntrega}
          onChange={(v) => setModoEntrega(v as typeof modoEntrega)}
        />
      </div>

      {/* Seção 5 — Inputs de valores */}
      <div className="space-y-6 pt-4 border-t border-[var(--color-border)]">
        <CurrencyInput
          label="Valor p/ m²"
          value={valorM2}
          onChange={setValorM2}
          placeholder="0,00"
        />

        <MedidasAlturaLarguraInput
          label="Medidas da Porta (Altura × Largura)"
          altura={alturaPorta}
          largura={larguraPorta}
          onAlturaChange={setAlturaPorta}
          onLarguraChange={setLarguraPorta}
          id="medidas-porta"
          error={formTouched && !okMedidasPorta ? 'Informe altura e largura da porta.' : undefined}
        />

        {bandeirola && (
          <MedidasAlturaLarguraInput
            label="Medidas da Bandeirola (Altura × Largura)"
            altura={alturaBandeirola}
            largura={larguraBandeirola}
            onAlturaChange={setAlturaBandeirola}
            onLarguraChange={setLarguraBandeirola}
            id="medidas-bandeirola"
            error={formTouched && !okMedidasBandeirola ? 'Informe altura e largura da bandeirola.' : undefined}
          />
        )}

        {alizar && (
          <div>
            <label htmlFor="medida-alizar" className="text-sm font-medium text-white block mb-1.5">
              Medida do Alizar
            </label>
            <input
              id="medida-alizar"
              type="text"
              inputMode="decimal"
              value={medidaAlizar}
              onChange={(e) => setMedidaAlizar(e.target.value)}
              placeholder="Ex.: 0,10"
              className="w-full max-w-[200px] py-3 px-4 rounded-[var(--radius)] border bg-[var(--color-surface)] text-white placeholder-[var(--color-text-muted)] focus:outline-none focus:border-[var(--color-accent)] focus:ring-2 focus:ring-[var(--color-accent)] focus:ring-opacity-30"
            />
            <p className="mt-1 text-xs text-[var(--color-text-muted)]">Ex.: 0,10 → 0,10m</p>
            {formTouched && !okMedidaAlizar && (
              <p className="text-xs text-amber-400">Informe a medida do alizar.</p>
            )}
          </div>
        )}

        <CurrencyInput
          id="custo-deslocamento-porta"
          label="Custo de Deslocamento"
          value={custoDeslocamento}
          onChange={setCustoDeslocamento}
          placeholder="0,00"
          suffix={
            <button
              type="button"
              onClick={() => setCustoDeslocamento('0')}
              className="shrink-0 self-stretch flex items-center px-4 rounded-[var(--radius)] border border-[var(--color-border)] bg-[var(--color-surface)] text-[var(--color-text-muted)] hover:bg-[var(--color-surface-hover)] hover:text-[var(--color-text)] focus:outline-none focus:ring-2 focus:ring-[var(--color-accent)] text-sm font-medium transition-colors"
            >
              Sem custo
            </button>
          }
        />

        <div>
          <label htmlFor="descricao-adicional-porta" className="text-sm font-medium text-white block mb-1.5">
            Descrição Adicional <span className="text-[var(--color-text-muted)] font-normal">(opcional)</span>
          </label>
          <textarea
            id="descricao-adicional-porta"
            value={descricaoAdicional}
            onChange={(e) => setDescricaoAdicional(e.target.value)}
            placeholder="Texto adicional para o orçamento"
            className="w-full py-3 px-4 rounded-[var(--radius)] border border-[var(--color-border)] bg-[var(--color-surface)] text-white placeholder-[var(--color-text-muted)] focus:outline-none focus:ring-2 focus:ring-[var(--color-accent)] focus:border-transparent resize-y min-h-[80px]"
            rows={3}
          />
        </div>
      </div>

      <button
        type="button"
        onClick={handleSubmit}
        disabled={!canSubmit}
        className="mt-8 w-full py-4 rounded-[var(--radius-lg)] font-medium bg-[var(--color-accent)] text-[#0d0d0d] hover:bg-[var(--color-accent-hover)] focus:outline-none focus:ring-2 focus:ring-[var(--color-accent)] focus:ring-offset-2 focus:ring-offset-[var(--color-bg)] disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:bg-[var(--color-accent)] transition-all"
      >
        Avançar
      </button>
      <p className="text-sm text-[var(--color-text-muted)] mt-4">
        Na próxima tela você informará os dados do cliente.
      </p>
    </div>
  );
}
