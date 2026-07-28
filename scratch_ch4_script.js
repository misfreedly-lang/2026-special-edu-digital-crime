
    // ==========================================
    // PRESENTATION & DUAL-SCREEN ENGINE SCRIPT V3
    // ==========================================
    const slides = document.querySelectorAll('.slide');
    const totalSlides = slides.length;
    let currentIndex = 0;
    
    const slideNumElem = document.getElementById('current-slide-num');
    const notesPanel = document.getElementById('notes-panel');
    const notesContent = document.getElementById('notes-content');
    
    const indPart1 = document.getElementById('ind-part1');
    const indChap1 = document.getElementById('ind-chap1');
    const indChap2 = document.getElementById('ind-chap2');
    const indChap3 = document.getElementById('ind-chap3');
    const indChap4 = document.getElementById('ind-chap4');
    const indChap5 = document.getElementById('ind-chap5');

    // Detect if running inside an iframe (Presenter View preview mode)
    const isInsideIframe = (window !== window.top);

    // BroadcastChannel for Dual Screen Sync — ONLY for the main top-level window
    const channel = isInsideIframe ? null : new BroadcastChannel('presentation_sync_ch4');

    function updateSlide(index, fromSync = false) {
      if (index < 0) index = 0;
      if (index >= totalSlides) index = totalSlides - 1;
      
      slides.forEach((s, i) => {
        if (i === index) {
          s.classList.add('active');
        } else {
          s.classList.remove('active');
        }
      });

      currentIndex = index;

      // Update slide number display (padded 2 digits)
      slideNumElem.textContent = String(currentIndex + 1).padStart(2, '0');

      // Update presenter notes
      const currentNotes = slides[currentIndex].getAttribute('data-notes') || '발표자 노트가 없습니다.';
      notesContent.textContent = currentNotes;

      // Hide header and footer inside iframe preview mode
      if (isInsideIframe) {
        const hdr = document.querySelector('header');
        const ftr = document.querySelector('footer');
        if (hdr) hdr.style.display = 'none';
        if (ftr) ftr.style.display = 'none';
      }

      // Broadcast to Presenter Window if updated locally (only from main window)
      if (!fromSync && channel) {
        const nextNotes = (currentIndex < totalSlides - 1) ? slides[currentIndex + 1].getAttribute('data-notes') : '마지막 슬라이드입니다.';
        channel.postMessage({
          type: 'SLIDE_CHANGE',
          index: currentIndex,
          total: totalSlides,
          currentNotes: currentNotes,
          nextNotes: nextNotes
        });
      }

      // Sync Presenter Window UI (Preview & Script Editor)
      syncPresenterView();
    }

    function nextSlide() {
      if (currentIndex < totalSlides - 1) {
        updateSlide(currentIndex + 1);
      }
    }

    function prevSlide() {
      if (currentIndex > 0) {
        updateSlide(currentIndex - 1);
      }
    }

    function toggleNotes() {
      notesPanel.classList.toggle('visible');
    }

    function toggleFullscreen() {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen();
      } else {
        if (document.exitFullscreen) {
          document.exitFullscreen();
        }
      }
    }

    // Presenter View Window Launcher (Dual Screen Setup with Iframe Rendering)
    let presenterWindow = null;

    function getSlideNotes(idx) {
      if (idx < 0 || idx >= totalSlides) return '슬라이드가 없습니다.';
      return slides[idx].getAttribute('data-notes') || '발표자 노트가 없습니다.';
    }

    // Dedicated preview renderer for Presenter View Iframes (called via postMessage)
    function setSlideForPreview(index) {
      if (index < 0 || index >= totalSlides) return;
      slides.forEach((s, i) => {
        if (i === index) {
          s.classList.add('active');
        } else {
          s.classList.remove('active');
        }
      });
    }

    // Iframe listens for postMessage commands from parent/opener
    if (isInsideIframe) {
      window.addEventListener('message', (event) => {
        if (event.data && event.data.type === 'SET_SLIDE') {
          setSlideForPreview(event.data.index);
        }
      });
    }

    function syncPresenterView() {
      if (!presenterWindow || presenterWindow.closed) return;

      const doc = presenterWindow.document;
      const currNotes = getSlideNotes(currentIndex);

      // 1. Update slide numbers
      if (doc.getElementById('pv-slide-curr')) {
        doc.getElementById('pv-slide-curr').textContent = String(currentIndex + 1).padStart(2, '0');
        doc.getElementById('pv-slide-total').textContent = String(totalSlides).padStart(2, '0');
      }

      // 2. Update script editor
      if (doc.getElementById('pv-script-edit') && doc.activeElement !== doc.getElementById('pv-script-edit')) {
        doc.getElementById('pv-script-edit').value = currNotes;
      }

      // 3. Update Current Slide Iframe via postMessage (currentIndex)
      const iframeCurr = doc.getElementById('pv-iframe-curr');
      if (iframeCurr && iframeCurr.contentWindow) {
        iframeCurr.contentWindow.postMessage({ type: 'SET_SLIDE', index: currentIndex }, '*');
      }

      // 4. Update Next Slide Iframe via postMessage (currentIndex + 1)
      const iframeNext = doc.getElementById('pv-iframe-next');
      if (iframeNext && iframeNext.contentWindow) {
        if (currentIndex < totalSlides - 1) {
          iframeNext.style.display = 'block';
          iframeNext.contentWindow.postMessage({ type: 'SET_SLIDE', index: currentIndex + 1 }, '*');
          if (doc.getElementById('pv-next-title')) doc.getElementById('pv-next-title').textContent = `🔮 NEXT (SLIDE ${String(currentIndex + 2).padStart(2, '0')})`;
        } else {
          iframeNext.style.display = 'none';
          if (doc.getElementById('pv-next-title')) doc.getElementById('pv-next-title').textContent = '🔮 NEXT (마지막 슬라이드)';
        }
      }
    }

    function openPresenterView() {
      const windowFeatures = 'width=1400,height=850,menubar=no,toolbar=no,location=no,status=no';
      presenterWindow = window.open('', 'PresenterViewWindow_CH4', windowFeatures);

      if (presenterWindow) {
        presenterWindow.document.write(`
          <!DOCTYPE html>
          <html lang="ko">
          <head>
            <meta charset="UTF-8">
            <title>강사 발표자 모드 v3 (CHAPTER Ⅳ)</title>
            <style>
              * { box-sizing: border-box; margin: 0; padding: 0; }
              body { background: #070A10; color: #EAEEF3; font-family: 'Noto Sans KR', -apple-system, sans-serif; padding: 14px; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
              
              .pv-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #FFB733; padding-bottom: 8px; margin-bottom: 10px; }
              .pv-title { font-size: 1.1rem; font-weight: 800; color: #FFB733; display: flex; align-items: center; gap: 8px; }
              .pv-timer { font-family: monospace; font-size: 1.3rem; color: #36CFC9; font-weight: bold; background: rgba(54,207,201,0.1); padding: 4px 12px; border-radius: 6px; border: 1px solid rgba(54,207,201,0.3); }

              /* 50:50 Layout Split */
              .pv-main-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; flex: 1; min-height: 0; }
              
              /* Left Column (50%): Previews (Top Large Current, Bottom Next) */
              .pv-left-col { display: flex; flex-direction: column; gap: 10px; height: 100%; overflow: hidden; }
              
              .preview-card { background: #121824; border: 1px solid rgba(255,183,51,0.3); border-radius: 8px; padding: 8px; display: flex; flex-direction: column; overflow: hidden; }
              .preview-card-header { font-size: 0.8rem; font-weight: 800; color: #FFB733; margin-bottom: 6px; display: flex; justify-content: space-between; }
              
              .iframe-wrap { flex: 1; background: #0D121D; border-radius: 6px; border: 1px solid rgba(255,255,255,0.15); overflow: hidden; position: relative; width: 100%; height: 100%; }
              .iframe-wrap iframe { width: 100%; height: 100%; border: none; }

              /* Right Column (50%): Large Script Editor */
              .pv-right-col { background: #121824; border: 1px solid rgba(255,183,51,0.3); border-radius: 8px; padding: 12px; display: flex; flex-direction: column; height: 100%; }
              .script-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
              .script-title { font-size: 0.95rem; font-weight: 800; color: #FFB733; }
              
              .script-editor { flex: 1; background: #080C14; color: #FFF; border: 1px solid rgba(255,183,51,0.4); border-radius: 8px; padding: 16px; font-size: 1.3rem; line-height: 1.8; font-family: inherit; resize: none; outline: none; }
              .script-editor:focus { border-color: #FFB733; box-shadow: 0 0 12px rgba(255,183,51,0.4); }

              /* Bottom Controls Bar */
              .pv-footer { margin-top: 10px; display: flex; justify-content: space-between; align-items: center; }
              .pv-btn-group { display: flex; gap: 10px; }
              .btn-pv { background: #FFB733; color: #000; border: none; padding: 10px 24px; font-size: 1.05rem; font-weight: 800; border-radius: 6px; cursor: pointer; transition: all 0.2s ease; }
              .btn-pv:hover { background: #FFC857; transform: translateY(-2px); }
              .btn-pv-sub { background: rgba(255,255,255,0.1); color: #FFF; border: 1px solid rgba(255,255,255,0.2); }
              .btn-pv-sub:hover { background: rgba(255,255,255,0.2); }
            </style>
          </head>
          <body>
            <div class="pv-header">
              <div class="pv-title">
                <span>🎙️ 강사 발표자 모드 v3 (CHAPTER Ⅳ)</span>
                <span style="font-size: 0.8rem; color: #C4D1DF; font-weight: normal;">| 16:9 실물 Iframe 0초 실시간 동기화 중</span>
              </div>
              <div class="pv-timer" id="pv-clock">⏱ 진행시간: 00:00</div>
            </div>

            <div class="pv-main-grid">
              <!-- Left Column: Previews (50% Width) -->
              <div class="pv-left-col">
                <!-- Current Slide (Top Large 65% Height) -->
                <div class="preview-card" style="height: 65%;">
                  <div class="preview-card-header">
                    <span>📺 현재 송출 화면 (대형 실물 16:9)</span>
                    <span style="color: #36CFC9;">SLIDE <span id="pv-slide-curr">01</span> / <span id="pv-slide-total">18</span></span>
                  </div>
                  <div class="iframe-wrap">
                    <iframe id="pv-iframe-curr" src="강의_발표자료_챕터4.html"></iframe>
                  </div>
                </div>

                <!-- Next Slide (Bottom 35% Height) -->
                <div class="preview-card" style="height: 35%;">
                  <div class="preview-card-header">
                    <span id="pv-next-title">🔮 NEXT (다음 슬라이드)</span>
                  </div>
                  <div class="iframe-wrap">
                    <iframe id="pv-iframe-next" src="강의_발표자료_챕터4.html" style="opacity: 0.85;"></iframe>
                  </div>
                </div>
              </div>

              <!-- Right Column: Large Script Editor (50% Width) -->
              <div class="pv-right-col">
                <div class="script-header">
                  <div class="script-title">🎙️ 강사 원본 스크립트 (실시간 수정 및 영구 자동 저장 💾)</div>
                  <span style="font-size: 0.75rem; color: #36CFC9; font-weight: bold;">* 수정 즉시 내 PC 브라우저(localStorage)에 영구 보존됩니다.</span>
                </div>
                <textarea class="script-editor" id="pv-script-edit" placeholder="스크립트를 입력하거나 수정하세요..."></textarea>
              </div>
            </div>

            <!-- Footer Controls -->
            <div class="pv-footer">
              <div style="font-size: 0.85rem; color: #C4D1DF;">
                키보드 <kbd>◀</kbd> <kbd>▶</kbd> 또는 <kbd>Space</kbd>로 발표 제어 가능
              </div>

              <div class="pv-btn-group">
                <button class="btn-pv btn-pv-sub" onclick="window.opener.prevSlide()">◀ 이전 슬라이드</button>
                <button class="btn-pv" onclick="window.opener.nextSlide()">다음 슬라이드 ▶</button>
              </div>
            </div>

            <script>
              // Timer Logic
              let seconds = 0;
              setInterval(() => {
                seconds++;
                const m = String(Math.floor(seconds / 60)).padStart(2, '0');
                const s = String(seconds % 60).padStart(2, '0');
                document.getElementById('pv-clock').textContent = '⏱ 진행시간: ' + m + ':' + s;
              }, 1000);

              // Live Script Edit Sync back to Parent Window
              const scriptEditor = document.getElementById('pv-script-edit');
              scriptEditor.addEventListener('input', (e) => {
                if (window.opener && !window.opener.closed) {
                  window.opener.updateCurrentScriptNotes(e.target.value);
                }
              });

              // Keyboard Navigation Passthrough to Main Window
              document.addEventListener('keydown', (e) => {
                if (document.activeElement === scriptEditor) return;
                if (e.key === 'ArrowRight' || e.key === 'Space' || e.key === 'PageDown') {
                  window.opener.nextSlide();
                  e.preventDefault();
                } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
                  window.opener.prevSlide();
                  e.preventDefault();
                }
              });
            <\/script>
          </body>
          </html>
        `);
        presenterWindow.document.close();

        setTimeout(() => {
          syncPresenterView();
          const iframeCurr = presenterWindow.document.getElementById('pv-iframe-curr');
          const iframeNext = presenterWindow.document.getElementById('pv-iframe-next');
          if (iframeCurr) iframeCurr.onload = () => syncPresenterView();
          if (iframeNext) iframeNext.onload = () => syncPresenterView();
        }, 300);
      }
    }

    // Load saved script notes from localStorage if available
    slides.forEach((s, i) => {
      const savedNotes = localStorage.getItem(`script_notes_ch4_slide_${i}`);
      if (savedNotes !== null) {
        s.setAttribute('data-notes', savedNotes);
      }
    });

    // Function to handle live script modification from presenter window
    function updateCurrentScriptNotes(newNotes) {
      if (slides[currentIndex]) {
        slides[currentIndex].setAttribute('data-notes', newNotes);
        notesContent.textContent = newNotes;
        // Save to Browser's LocalStorage permanently
        localStorage.setItem(`script_notes_ch4_slide_${currentIndex}`, newNotes);
      }
    }

    // === Only activate interactive features in the MAIN top-level window ===
    if (!isInsideIframe) {
      channel.onmessage = (event) => {
        if (event.data.type === 'SLIDE_CHANGE') {
          const { index } = event.data;
          if (currentIndex !== index) {
            updateSlide(index, true);
          }
        }
      };

      document.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowRight' || e.key === 'Space' || e.key === 'PageDown') {
          nextSlide();
        } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
          prevSlide();
        } else if (e.key === 'n' || e.key === 'N') {
          toggleNotes();
        } else if (e.key === 'f' || e.key === 'F') {
          toggleFullscreen();
        } else if (e.key === 'p' || e.key === 'P') {
          openPresenterView();
        }
      });

      updateSlide(0);
    } else {
      const hdr = document.querySelector('header');
      const ftr = document.querySelector('footer');
      if (hdr) hdr.style.display = 'none';
      if (ftr) ftr.style.display = 'none';
    }
  