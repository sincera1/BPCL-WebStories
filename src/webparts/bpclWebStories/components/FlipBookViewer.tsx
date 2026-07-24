import * as React from 'react';
import HTMLFlipBook from "react-pageflip";
import styles from './BpclWebStories.module.scss';

import * as pdfjsLib from "pdfjs-dist/build/pdf";
import "pdfjs-dist/build/pdf.worker.entry";
// new code
import {IconButton} from "@fluentui/react";

import type { RenderParameters } from "pdfjs-dist/types/src/display/api";

interface IProps {
  fileUrl: string;
  onClose: () => void;
}

const FlipBookViewer: React.FC<IProps> = ({ fileUrl, onClose }) => {

// const FlipBook = HTMLFlipBook as any; 
const FlipBook = HTMLFlipBook as unknown as React.ComponentType<Record<string, unknown>>;
const [pages, setPages] = React.useState<string[]>([]);
const [zoom, setZoom] = React.useState(1);
const [isFullScreen, setIsFullScreen] = React.useState(false);
const containerRef = React.useRef<HTMLDivElement>(null);

// new code
const [currentPage, setCurrentPage] = React.useState(1);
// const flipBookRef = React.useRef<any>(null);
const flipBookRef = React.useRef<{
  pageFlip: () => {
    flipNext: () => void;
    flipPrev: () => void;
  };
} | null>(null);

const zoomIn = (): void  => setZoom(prev => Math.min(prev + 0.2, 2));
const zoomOut = (): void => setZoom(prev => Math.max(prev - 0.2, 0.6));
// const toggleFullScreen = () => {
//   if (!document.fullscreenElement) {
//     containerRef.current?.requestFullscreen();
//     setIsFullScreen(true);
//   } else {
//     document.exitFullscreen();
//     setIsFullScreen(false);
//   }
// };
const toggleFullScreen = (): void => {
  if (!document.fullscreenElement) {
  containerRef.current
      ?.requestFullscreen()
      .then(() => {
        setIsFullScreen(true);
      })
      .catch((error) => {
        console.error("Failed to enter fullscreen:", error);
      });
  } else {
    document.exitFullscreen().catch((error) => {
      console.error("Error loading fullscreen:", error);
      setIsFullScreen(false);
    });
  }
};
// new code
const nextPage = (): void => {
  flipBookRef.current?.pageFlip()?.flipNext();
};

const prevPage = (): void => {
  flipBookRef.current?.pageFlip()?.flipPrev();
};
//
const loadPdf = async (): Promise<void> => {
  try {
    console.log("Loading PDF:", fileUrl);

    // ✅ Fetch file using SharePoint auth
    const response = await fetch(fileUrl, {
      headers: {
        "Accept": "application/pdf"
      },
      credentials: "include"
    });

    const blob = await response.blob();

    // ✅ Convert blob to URL
    const blobUrl = URL.createObjectURL(blob);

    const pdf = await pdfjsLib.getDocument(blobUrl).promise;

    const tempPages: string[] = [];

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale: 1.5 });

      const canvas = document.createElement("canvas");
      const context = canvas.getContext("2d");

      canvas.height = viewport.height;
      canvas.width = viewport.width;

      // await page.render({
      //   canvasContext: context!,
      //   viewport,
      //   canvas
      // } as any).promise;
       const renderContext: RenderParameters = {
          canvasContext: context!,
          viewport
        };
      await page.render(renderContext).promise;

      tempPages.push(canvas.toDataURL());
    }

    setPages(tempPages);

  } catch (error) {
    console.error("PDF Load Error:", error);
  }
};

 React.useEffect(() => {
  loadPdf().catch((error) => {
    console.error("Error loading PDF:", error);
  });
}, [fileUrl]);


React.useEffect(() => {

  const handleKey = (e: KeyboardEvent): void => {

    if (e.key === "ArrowRight") {
      nextPage();
    }

    if (e.key === "ArrowLeft") {
      prevPage();
    }

  };

  window.addEventListener("keydown", handleKey);

  return () => {
    window.removeEventListener("keydown", handleKey);
  };

}, []);

 return (

  <div ref={containerRef} className={styles.viewerMain}>

    {/* 🔥 SINGLE TOOLBAR ONLY */}
    <div className={styles.toolbar}>

      {/* <IconButton  iconProps={{ iconName: "Back" }} title="Back" onClick={onClose} /> */}
       <button onClick={onClose}>⬅ Back</button>

      {/* <IconButton  iconProps={{ iconName: "ChevronLeft" }}  title="Previous" onClick={prevPage} /> */}
       <button onClick={prevPage}>⏮ Prev</button>

     {/* <div className={styles.pageInfo}>
        {pages.length > 0
          ? `Page ${currentPage} / ${pages.length}`
          : "Loading PDF..."
        }
      </div> */}

      <div className={styles.pageInfo}>
  {pages.length > 0
    ? (() => {

        // Cover page
        if (currentPage === 1) {
          return `Page 1 / ${pages.length}`;
        }

        const rightPage = Math.min(
          currentPage + 1,
          pages.length
        );

        return `Pages ${currentPage}-${rightPage} / ${pages.length}`;

      })()
    : "Loading PDF..."
  }
</div>

      {/* <IconButton iconProps={{ iconName: "ChevronRight" }}  title="Next"  onClick={nextPage}  /> */}
      <button onClick={nextPage}>Next ⏭</button>

      <IconButton  iconProps={{ iconName: "ZoomOut" }}  title="Zoom Out"  onClick={zoomOut} />

      <IconButton  iconProps={{ iconName: "ZoomIn" }}  title="Zoom In"  onClick={zoomIn} />

      <IconButton iconProps={{
          iconName: isFullScreen
            ? "BackToWindow"
            : "FullScreen"
        }}
        title="Fullscreen"
        onClick={toggleFullScreen}
      />

    </div>

    {/* 🔥 FLIPBOOK */}
    <div className={styles.flipbookCenter}>

      <div
        className={styles.flipbookWrapper}
        style={{ transform: `scale(${zoom})` }}
      >

        <FlipBook
          ref={flipBookRef}
          width={800}
          height={1000}
          size="stretch"
          minWidth={500}
          maxWidth={1200}
          minHeight={600}
          maxHeight={1400}
          usePortrait={true}
          showCover={true}
          mobileScrollSupport={true}
          clickEventForward={false}
          onFlip={(e: { data: number }) =>
            setCurrentPage(e.data + 1)
          }
        >

          {pages.map((img, i) => (
            <div key={i} className={styles.page}>
              <img
                src={img}
                className={styles.pageImage}
              />
            </div>
          ))}

        </FlipBook>

      </div>

    </div>

  </div>

);
  
};

export default FlipBookViewer;