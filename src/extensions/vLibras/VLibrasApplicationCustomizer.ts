import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IVlibrasApplicationCustomizerProperties {
  /**
   * Se true, carrega o script de um arquivo local (Site Assets).
   * Se false (padrão), carrega direto de vlibras.gov.br.
   */
  useLocalScript?: boolean;

  /**
   * URL do script local. Usada apenas quando useLocalScript = true.
   */
  localScriptUrl?: string;
}

const LOG_SOURCE = 'VlibrasApplicationCustomizer';
const VLIBRAS_SCRIPT_REMOTE = 'https://vlibras.gov.br/app/vlibras-plugin.js';
const VLIBRAS_ROOT_PATH = 'https://vlibras.gov.br/app';

export default class VlibrasApplicationCustomizer
  extends BaseApplicationCustomizer<IVlibrasApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent | undefined;

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Application Customizer iniciado');

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholders);
    this._renderPlaceholders();

    return Promise.resolve();
  }

  private _renderPlaceholders(): void {
    if (this._bottomPlaceholder) {
      return;
    }

    this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Bottom,
      { onDispose: this._onDispose }
    );

    if (!this._bottomPlaceholder || !this._bottomPlaceholder.domElement) {
      Log.warn(LOG_SOURCE, 'Placeholder Bottom indisponível.');
      return;
    }

    // HTML oficial do widget VLibras
    this._bottomPlaceholder.domElement.innerHTML = `
      <div vw class="enabled">
        <div vw-access-button class="active"></div>
        <div vw-plugin-wrapper>
          <div class="vw-plugin-top-wrapper"></div>
        </div>
      </div>
    `;

    this._loadVlibrasScript()
      .then(() => this._initWidget())
      .catch((err) => {
        Log.error(LOG_SOURCE, new Error(String(err)));
      });
  }

  private async _loadVlibrasScript(): Promise<void> {
    const useLocal = this.properties.useLocalScript === true;
    const localUrl = this.properties.localScriptUrl;

    if ((window as any).VLibras && (window as any).VLibras.Widget) {
      Log.info(LOG_SOURCE, 'VLibras já estava carregado.');
      return;
    }

    if (useLocal && localUrl) {
      Log.info(LOG_SOURCE, `Carregando script LOCAL: ${localUrl}`);
      try {
        await SPComponentLoader.loadScript(localUrl, {
          globalExportsName: 'VLibras'
        });
        return;
      } catch (err) {
        Log.warn(LOG_SOURCE, `Falha no script local: ${err}. Caindo para o remoto.`);
      }
    }

    Log.info(LOG_SOURCE, `Carregando script REMOTO: ${VLIBRAS_SCRIPT_REMOTE}`);
    await SPComponentLoader.loadScript(VLIBRAS_SCRIPT_REMOTE, {
      globalExportsName: 'VLibras'
    });
  }

  private _initWidget(): void {
    const vlibras = (window as any).VLibras;
    if (!vlibras || !vlibras.Widget) {
      Log.error(LOG_SOURCE, new Error('VLibras.Widget não está disponível.'));
      return;
    }

    Log.info(LOG_SOURCE, 'Inicializando widget VLibras...');
    // eslint-disable-next-line no-new
    new vlibras.Widget(VLIBRAS_ROOT_PATH);
    Log.info(LOG_SOURCE, 'Widget VLibras inicializado.');

    this._adjustWidgetPosition();
  }

  /**
   * O widget do VLibras renderiza dentro de um shadow DOM, então CSS externo
   * não o alcança. Este método injeta um ajuste de posição dentro do shadow root.
   */
  private _adjustWidgetPosition(): void {
    let attempts = 0;
    const maxAttempts = 20;

    const tryInject = () => {
      attempts++;

      const host = document.querySelector('#vlibras-access-wrapper') as HTMLElement | null;
      const shadow = host && (host as any).shadowRoot as ShadowRoot | null;

      if (shadow) {
        if (!shadow.querySelector('#vlibras-custom-style')) {
          const style = document.createElement('style');
          style.id = 'vlibras-custom-style';
          style.textContent = `
            #vlibras-access {
              right: 72px !important;
             bottom: 18px !important;
            }
          `;
          shadow.appendChild(style);
          Log.info(LOG_SOURCE, 'Ajuste de posição aplicado.');
        }
        return;
      }

      if (attempts < maxAttempts) {
        setTimeout(tryInject, 500);
      } else {
        Log.warn(LOG_SOURCE, 'Não foi possível encontrar o shadow root do VLibras.');
      }
    };

    tryInject();
  }

  private _onDispose(): void {
    Log.info(LOG_SOURCE, 'Application Customizer descartado.');
  }
}
