// Filmstrip navigator: slide thumbnails + keyboard navigation.

import type { PresentationViewModel } from "./types.ts";
import { SlideRenderer } from "./renderer.ts";

export type SlideChangeCallback = (index: number) => void;

export class Navigator {
  private filmstrip: HTMLElement;
  private renderer: SlideRenderer;
  private model: PresentationViewModel | null = null;
  private currentSlide = 0;
  private onChange: SlideChangeCallback | null = null;
  private keyHandler: ((e: KeyboardEvent) => void) | null = null;

  constructor(filmstrip: HTMLElement, renderer: SlideRenderer) {
    this.filmstrip = filmstrip;
    this.renderer = renderer;
  }

  setModel(model: PresentationViewModel): void {
    this.model = model;
    this.currentSlide = 0;
    this.buildFilmstrip();
    this.highlightActive();
  }

  setOnChange(cb: SlideChangeCallback): void {
    this.onChange = cb;
  }

  goToSlide(index: number): void {
    if (!this.model) return;
    const clamped = Math.max(0, Math.min(index, this.model.slides.length - 1));
    this.currentSlide = clamped;
    this.renderer.goToSlide(clamped);
    this.highlightActive();
    this.onChange?.(clamped);
  }

  nextSlide(): void {
    this.goToSlide(this.currentSlide + 1);
  }

  prevSlide(): void {
    this.goToSlide(this.currentSlide - 1);
  }

  slideCount(): number {
    return this.model?.slides.length ?? 0;
  }

  getCurrentSlide(): number {
    return this.currentSlide;
  }

  /** Attach keyboard listeners to the given element. */
  attachKeyboard(target: HTMLElement | Window): void {
    this.keyHandler = (e: KeyboardEvent) => {
      switch (e.key) {
        case "ArrowRight":
        case "ArrowDown":
        case "PageDown":
          e.preventDefault();
          this.nextSlide();
          break;
        case "ArrowLeft":
        case "ArrowUp":
        case "PageUp":
          e.preventDefault();
          this.prevSlide();
          break;
        case "Home":
          e.preventDefault();
          this.goToSlide(0);
          break;
        case "End":
          e.preventDefault();
          if (this.model) {
            this.goToSlide(this.model.slides.length - 1);
          }
          break;
      }
    };
    target.addEventListener("keydown", this.keyHandler as EventListener);
  }

  /** Remove keyboard listeners. */
  detachKeyboard(target: HTMLElement | Window): void {
    if (this.keyHandler) {
      target.removeEventListener("keydown", this.keyHandler as EventListener);
      this.keyHandler = null;
    }
  }

  destroy(): void {
    this.filmstrip.innerHTML = "";
  }

  // ---------- Filmstrip ----------

  private buildFilmstrip(): void {
    this.filmstrip.innerHTML = "";
    if (!this.model) return;

    // Thumbnail width: filmstrip is 160px, minus padding = ~140px usable.
    const thumbnailWidth = 136;

    for (let i = 0; i < this.model.slides.length; i++) {
      const wrapper = document.createElement("div");
      wrapper.className = "pptview-thumbnail-wrapper";
      wrapper.dataset["index"] = String(i);

      // Slide number label
      const numLabel = document.createElement("div");
      numLabel.className = "pptview-thumbnail-number";
      numLabel.textContent = String(i + 1);
      wrapper.appendChild(numLabel);

      // Render thumbnail
      const thumb = this.renderer.renderThumbnail(i, thumbnailWidth);
      wrapper.appendChild(thumb);

      wrapper.addEventListener("click", () => {
        this.goToSlide(i);
      });

      this.filmstrip.appendChild(wrapper);
    }
  }

  private highlightActive(): void {
    const wrappers = this.filmstrip.querySelectorAll(
      ".pptview-thumbnail-wrapper",
    );
    wrappers.forEach((el, idx) => {
      el.classList.toggle("active", idx === this.currentSlide);
    });

    // Scroll active thumbnail into view
    const active = this.filmstrip.querySelector(
      ".pptview-thumbnail-wrapper.active",
    );
    active?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }
}
