type ArtProps = {
  className?: string;
};

export function PortalBackdropArt({ className = "" }: ArtProps) {
  return (
    <svg
      aria-hidden="true"
      className={className}
      viewBox="0 0 960 720"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <defs>
        <linearGradient
          id="portal-bg"
          x1="160"
          y1="40"
          x2="820"
          y2="680"
          gradientUnits="userSpaceOnUse"
        >
          <stop stopColor="#FFF9EF" />
          <stop offset="0.45" stopColor="#EAF1F8" />
          <stop offset="1" stopColor="#DDECF1" />
        </linearGradient>
        <linearGradient
          id="portal-frame"
          x1="180"
          y1="100"
          x2="760"
          y2="620"
          gradientUnits="userSpaceOnUse"
        >
          <stop stopColor="#0F766E" />
          <stop offset="0.5" stopColor="#1E88E5" />
          <stop offset="1" stopColor="#C0841A" />
        </linearGradient>
        <linearGradient
          id="portal-card"
          x1="240"
          y1="160"
          x2="660"
          y2="560"
          gradientUnits="userSpaceOnUse"
        >
          <stop stopColor="#FFFFFF" />
          <stop offset="1" stopColor="#F4F7FB" />
        </linearGradient>
        <linearGradient id="portal-glow" x1="0" y1="0" x2="1" y2="1">
          <stop stopColor="#14B8A6" stopOpacity="0.9" />
          <stop offset="1" stopColor="#F59E0B" stopOpacity="0.95" />
        </linearGradient>
        <filter id="portal-shadow" x="-20%" y="-20%" width="140%" height="140%">
          <feDropShadow dx="0" dy="16" stdDeviation="16" floodColor="#082032" floodOpacity="0.18" />
        </filter>
        <filter id="portal-soft" x="-40%" y="-40%" width="180%" height="180%">
          <feGaussianBlur stdDeviation="18" />
        </filter>
      </defs>

      <rect x="12" y="12" width="936" height="696" rx="42" fill="url(#portal-bg)" />

      <g className="art-float art-float-slow" opacity="0.9" filter="url(#portal-soft)">
        <circle cx="180" cy="154" r="82" fill="#0F766E" fillOpacity="0.18" />
        <circle cx="768" cy="160" r="92" fill="#1E88E5" fillOpacity="0.12" />
        <circle cx="740" cy="560" r="120" fill="#C0841A" fillOpacity="0.12" />
      </g>

      <g className="art-float art-float-reverse" filter="url(#portal-shadow)">
        <path
          d="M114 170C114 147.909 131.909 130 154 130H402C424.091 130 442 147.909 442 170V548C442 570.091 424.091 588 402 588H154C131.909 588 114 570.091 114 548V170Z"
          fill="url(#portal-card)"
          stroke="rgba(255,255,255,0.9)"
          strokeWidth="3"
        />
        <path
          d="M138 176C138 160.536 150.536 148 166 148H228C243.464 148 256 160.536 256 176V196C256 211.464 243.464 224 228 224H166C150.536 224 138 211.464 138 196V176Z"
          fill="url(#portal-glow)"
          opacity="0.85"
        />
        <circle cx="170" cy="307" r="30" fill="#0F766E" fillOpacity="0.12" />
        <circle cx="260" cy="307" r="30" fill="#1E88E5" fillOpacity="0.12" />
        <circle cx="350" cy="307" r="30" fill="#C0841A" fillOpacity="0.12" />
        <rect x="145" y="392" width="246" height="18" rx="9" fill="#D9E4EE" />
        <rect x="145" y="424" width="196" height="14" rx="7" fill="#E6EEF5" />
        <rect x="145" y="454" width="222" height="14" rx="7" fill="#E6EEF5" />
        <rect x="145" y="484" width="176" height="14" rx="7" fill="#E6EEF5" />
      </g>

      <g className="art-float" filter="url(#portal-shadow)">
        <path
          d="M520 124C520 103.013 537.013 86 558 86H822C842.987 86 860 103.013 860 124V572C860 592.987 842.987 610 822 610H558C537.013 610 520 592.987 520 572V124Z"
          fill="rgba(255,255,255,0.78)"
          stroke="rgba(255,255,255,0.9)"
          strokeWidth="3"
        />
        <path
          d="M553 154C553 140.745 563.745 130 577 130H804C817.255 130 828 140.745 828 154V200C828 213.255 817.255 224 804 224H577C563.745 224 553 213.255 553 200V154Z"
          fill="#0F766E"
          fillOpacity="0.1"
        />
        <rect x="590" y="272" width="176" height="22" rx="11" fill="#0F766E" fillOpacity="0.14" />
        <rect x="590" y="316" width="230" height="18" rx="9" fill="#1E88E5" fillOpacity="0.12" />
        <rect x="590" y="352" width="150" height="18" rx="9" fill="#C0841A" fillOpacity="0.16" />
        <rect x="590" y="390" width="204" height="18" rx="9" fill="#0F766E" fillOpacity="0.12" />
        <rect x="590" y="428" width="132" height="18" rx="9" fill="#1E88E5" fillOpacity="0.12" />
        <path
          d="M676 488C719.037 488 754 466.866 754 440.75C754 414.634 719.037 393.5 676 393.5C632.963 393.5 598 414.634 598 440.75C598 466.866 632.963 488 676 488Z"
          fill="#C0841A"
          fillOpacity="0.12"
        />
        <path
          d="M676 462C695.882 462 712 452.552 712 440.9C712 429.248 695.882 419.8 676 419.8C656.118 419.8 640 429.248 640 440.9C640 452.552 656.118 462 676 462Z"
          fill="#0F766E"
          fillOpacity="0.18"
        />
      </g>

      <g className="art-float art-float-delayed">
        <circle cx="470" cy="150" r="18" fill="#1E88E5" fillOpacity="0.2" />
        <circle cx="490" cy="176" r="8" fill="#C0841A" fillOpacity="0.3" />
        <circle cx="472" cy="520" r="12" fill="#0F766E" fillOpacity="0.18" />
        <path
          d="M488 108l7 15 15 7-15 7-7 15-7-15-15-7 15-7 7-15Z"
          fill="#F5B640"
          fillOpacity="0.75"
        />
        <path
          d="M88 286l8 14 14 8-14 8-8 14-8-14-14-8 14-8 8-14Z"
          fill="#0F766E"
          fillOpacity="0.55"
        />
        <path
          d="M864 240l8 14 14 8-14 8-8 14-8-14-14-8 14-8 8-14Z"
          fill="#1E88E5"
          fillOpacity="0.45"
        />
      </g>
    </svg>
  );
}

export function LoadingGlyphArt({ className = "" }: ArtProps) {
  return (
    <svg
      aria-hidden="true"
      className={className}
      viewBox="0 0 240 240"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <defs>
        <linearGradient
          id="loading-core"
          x1="56"
          y1="44"
          x2="188"
          y2="194"
          gradientUnits="userSpaceOnUse"
        >
          <stop stopColor="#0F766E" />
          <stop offset="0.5" stopColor="#1E88E5" />
          <stop offset="1" stopColor="#C0841A" />
        </linearGradient>
        <filter id="loading-shadow" x="-20%" y="-20%" width="140%" height="140%">
          <feDropShadow dx="0" dy="10" stdDeviation="10" floodColor="#082032" floodOpacity="0.2" />
        </filter>
      </defs>

      <g className="art-float art-float-slow" filter="url(#loading-shadow)">
        <circle cx="120" cy="120" r="92" fill="rgba(255,255,255,0.78)" />
        <circle
          cx="120"
          cy="120"
          r="82"
          fill="none"
          stroke="url(#loading-core)"
          strokeWidth="8"
          opacity="0.85"
        />
        <circle
          cx="120"
          cy="120"
          r="58"
          fill="none"
          stroke="#D6E1EE"
          strokeWidth="10"
          strokeDasharray="16 16"
        />
        <circle cx="120" cy="120" r="34" fill="url(#loading-core)" opacity="0.92" />
        <rect x="109" y="109" width="22" height="22" rx="6" fill="#FFFFFF" opacity="0.9" />
        <circle cx="120" cy="120" r="8" fill="#0F172A" opacity="0.25" />
      </g>

      <g className="art-float art-float-delayed">
        <circle cx="120" cy="24" r="8" fill="#F5B640" />
        <circle cx="214" cy="120" r="8" fill="#1E88E5" />
        <circle cx="120" cy="216" r="8" fill="#0F766E" />
        <circle cx="26" cy="120" r="8" fill="#C0841A" />
        <path
          d="M183 47l7 13 13 7-13 7-7 13-7-13-13-7 13-7 7-13Z"
          fill="#F5B640"
          fillOpacity="0.8"
        />
        <path
          d="M57 180l7 13 13 7-13 7-7 13-7-13-13-7 13-7 7-13Z"
          fill="#0F766E"
          fillOpacity="0.65"
        />
      </g>
    </svg>
  );
}
