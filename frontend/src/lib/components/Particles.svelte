<script lang="ts">
    import { onMount } from "svelte";

    let canvas: HTMLCanvasElement;

    onMount(() => {
        const ctx = canvas.getContext("2d")!;
        let animationId: number;
        let particles: {
            x: number;
            y: number;
            vx: number;
            vy: number;
            size: number;
            opacity: number;
            pulse: number;
        }[] = [];

        function resize() {
            canvas.width = window.innerWidth;
            canvas.height = window.innerHeight;
        }

        function createParticles() {
            particles = [];
            const count = Math.floor((canvas.width * canvas.height) / 15000);
            for (let i = 0; i < count; i++) {
                particles.push({
                    x: Math.random() * canvas.width,
                    y: Math.random() * canvas.height,
                    vx: (Math.random() - 0.5) * 0.3,
                    vy: (Math.random() - 0.5) * 0.2 - 0.1,
                    size: Math.random() * 2 + 0.5,
                    opacity: Math.random() * 0.4 + 0.1,
                    pulse: Math.random() * Math.PI * 2,
                });
            }
        }

        function draw() {
            ctx.clearRect(0, 0, canvas.width, canvas.height);

            for (const p of particles) {
                p.x += p.vx;
                p.y += p.vy;
                p.pulse += 0.01;

                // Wrap around edges
                if (p.x < -10) p.x = canvas.width + 10;
                if (p.x > canvas.width + 10) p.x = -10;
                if (p.y < -10) p.y = canvas.height + 10;
                if (p.y > canvas.height + 10) p.y = -10;

                const dynamicOpacity =
                    p.opacity * (0.6 + 0.4 * Math.sin(p.pulse));

                // Get theme color from CSS variable
                const computed = getComputedStyle(document.documentElement);
                const primaryColor =
                    computed.getPropertyValue("--color-primary-500").trim() ||
                    "#f43f5e";

                // Parse hex to rgb
                const r = parseInt(primaryColor.slice(1, 3), 16);
                const g = parseInt(primaryColor.slice(3, 5), 16);
                const b = parseInt(primaryColor.slice(5, 7), 16);

                ctx.beginPath();
                ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
                ctx.fillStyle = `rgba(${r}, ${g}, ${b}, ${dynamicOpacity})`;
                ctx.fill();

                // Glow effect on larger particles
                if (p.size > 1.5) {
                    ctx.beginPath();
                    ctx.arc(p.x, p.y, p.size * 3, 0, Math.PI * 2);
                    ctx.fillStyle = `rgba(${r}, ${g}, ${b}, ${dynamicOpacity * 0.15})`;
                    ctx.fill();
                }
            }

            // Draw subtle connection lines between nearby particles
            for (let i = 0; i < particles.length; i++) {
                for (let j = i + 1; j < particles.length; j++) {
                    const dx = particles[i].x - particles[j].x;
                    const dy = particles[i].y - particles[j].y;
                    const dist = Math.sqrt(dx * dx + dy * dy);
                    if (dist < 120) {
                        const computed = getComputedStyle(
                            document.documentElement,
                        );
                        const primaryColor =
                            computed
                                .getPropertyValue("--color-primary-500")
                                .trim() || "#f43f5e";
                        const r = parseInt(primaryColor.slice(1, 3), 16);
                        const g = parseInt(primaryColor.slice(3, 5), 16);
                        const b = parseInt(primaryColor.slice(5, 7), 16);

                        ctx.beginPath();
                        ctx.moveTo(particles[i].x, particles[i].y);
                        ctx.lineTo(particles[j].x, particles[j].y);
                        ctx.strokeStyle = `rgba(${r}, ${g}, ${b}, ${0.06 * (1 - dist / 120)})`;
                        ctx.lineWidth = 0.5;
                        ctx.stroke();
                    }
                }
            }

            animationId = requestAnimationFrame(draw);
        }

        resize();
        createParticles();
        draw();

        window.addEventListener("resize", () => {
            resize();
            createParticles();
        });

        return () => {
            cancelAnimationFrame(animationId);
        };
    });
</script>

<canvas
    bind:this={canvas}
    class="fixed inset-0 z-0 pointer-events-none"
    aria-hidden="true"
></canvas>
