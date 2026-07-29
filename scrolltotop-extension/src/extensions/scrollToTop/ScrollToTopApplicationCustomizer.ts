import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

const LOG_SOURCE: string = 'ScrollToTopApplicationCustomizer';

export interface IScrollToTopApplicationCustomizerProperties {}

export default class ScrollToTopApplicationCustomizer
  extends BaseApplicationCustomizer<IScrollToTopApplicationCustomizerProperties> {

  private _button: HTMLButtonElement | undefined;
  private _scrollHandler: (() => void) | undefined;

  public onInit(): Promise<void> {
  Log.info(LOG_SOURCE, 'Initialized ScrollToTop');
  this._renderButton();
  return Promise.resolve();
}

  private _renderButton(): void {
    const button = document.createElement('button');
    button.id = 'spfx-scroll-to-top';
    button.setAttribute('aria-label', 'Scroll to top');
    button.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#ffffff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="18 15 12 9 6 15"></polyline>
      </svg>
    `;

    button.style.position = 'fixed';
    button.style.bottom = '24px';
    button.style.right = '24px';
    button.style.width = '44px';
    button.style.height = '44px';
    button.style.borderRadius = '50%';
    button.style.border = 'none';
    button.style.backgroundColor = 'var(--themePrimary, #0078d4)';
    button.style.color = '#ffffff';
    button.style.cursor = 'pointer';
    button.style.display = 'none';
    button.style.alignItems = 'center';
    button.style.justifyContent = 'center';
    button.style.boxShadow = '0 2px 8px rgba(0, 0, 0, 0.25)';
    button.style.zIndex = '9999';
    button.style.transition = 'opacity 0.2s ease, background-color 0.15s ease';
    button.style.opacity = '0';

    button.addEventListener('mouseenter', () => {
      button.style.backgroundColor = 'var(--themeDarkAlt, #106ebe)';
    });
    button.addEventListener('mouseleave', () => {
      button.style.backgroundColor = 'var(--themePrimary, #0078d4)';
    });

    button.addEventListener('click', () => {
      const scrollContainer = document.querySelector('[data-automation-id="contentScrollRegion"]');
      if (scrollContainer) {
        scrollContainer.scrollTo({ top: 0, behavior: 'smooth' });
      }
    });

    document.body.appendChild(button);
    this._button = button;

    const attachScrollListener = (): void => {
      const scrollContainer = document.querySelector('[data-automation-id="contentScrollRegion"]');
      if (scrollContainer) {
        this._scrollHandler = (): void => {
          if (scrollContainer.scrollTop > 300) {
            button.style.display = 'flex';
            setTimeout(() => { button.style.opacity = '1'; }, 10);
          } else {
            button.style.opacity = '0';
            setTimeout(() => { button.style.display = 'none'; }, 200);
          }
        };
        scrollContainer.addEventListener('scroll', this._scrollHandler);
      } else {
        setTimeout(attachScrollListener, 500);
      }
    };

    attachScrollListener();
  }

  public onDispose(): void {
    const scrollContainer = document.querySelector('[data-automation-id="contentScrollRegion"]');
    if (this._scrollHandler && scrollContainer) {
      scrollContainer.removeEventListener('scroll', this._scrollHandler);
    }
    if (this._button) {
      this._button.remove();
    }
  }
}