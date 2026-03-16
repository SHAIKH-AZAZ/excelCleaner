"use client";

/**
 * Lightweight CSS-only animated background.
 * Replaces the heavy Three.js WebGL shader with performant CSS gradients
 * and animations that use GPU-composited properties (transform, opacity).
 */
const AnimatedShaderBackground = () => {
  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 0,
        pointerEvents: "none",
        overflow: "hidden",
        background: "linear-gradient(135deg, #0a0e1a 0%, #0d1117 40%, #0a0e1a 100%)",
      }}
    >
      {/* Soft aurora blobs – CSS only, GPU-composited */}
      <div
        style={{
          position: "absolute",
          top: "-30%",
          left: "-10%",
          width: "60%",
          height: "60%",
          borderRadius: "50%",
          background:
            "radial-gradient(circle, rgba(52,211,153,0.12) 0%, transparent 70%)",
          filter: "blur(80px)",
          animation: "bgFloat1 18s ease-in-out infinite alternate",
          willChange: "transform",
        }}
      />
      <div
        style={{
          position: "absolute",
          bottom: "-20%",
          right: "-10%",
          width: "55%",
          height: "55%",
          borderRadius: "50%",
          background:
            "radial-gradient(circle, rgba(99,102,241,0.10) 0%, transparent 70%)",
          filter: "blur(80px)",
          animation: "bgFloat2 22s ease-in-out infinite alternate",
          willChange: "transform",
        }}
      />
      <div
        style={{
          position: "absolute",
          top: "40%",
          left: "50%",
          width: "40%",
          height: "40%",
          borderRadius: "50%",
          background:
            "radial-gradient(circle, rgba(167,139,250,0.08) 0%, transparent 70%)",
          filter: "blur(90px)",
          animation: "bgFloat3 25s ease-in-out infinite alternate",
          willChange: "transform",
        }}
      />

      {/* Keyframes injected via <style> to keep it self-contained */}
      <style>{`
        @keyframes bgFloat1 {
          0%   { transform: translate(0, 0) scale(1); }
          100% { transform: translate(8%, 12%) scale(1.15); }
        }
        @keyframes bgFloat2 {
          0%   { transform: translate(0, 0) scale(1); }
          100% { transform: translate(-10%, -8%) scale(1.1); }
        }
        @keyframes bgFloat3 {
          0%   { transform: translate(0, 0) scale(1); opacity: 0.7; }
          100% { transform: translate(-6%, 10%) scale(1.2); opacity: 1; }
        }
      `}</style>
    </div>
  );
};

export default AnimatedShaderBackground;
