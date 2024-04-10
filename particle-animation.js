// Get the canvas element
const canvas = document.getElementById('particle-canvas');
const ctx = canvas.getContext('2d');

// Set canvas dimensions
canvas.width = window.innerWidth;
canvas.height = window.innerHeight;

// Array to hold particles
const particles = [];

// Function to generate a particle
function createParticle() {
  const particle = {
    x: Math.random() * canvas.width,
    y: Math.random() * canvas.height,
    size: Math.random() * 1.6, // Random size between 1 and 4
    speedX: Math.random() * 3 - 1.5, // Random horizontal speed
    speedY: Math.random() * 3 - 1.5, // Random vertical speed
  };
  particles.push(particle);
}

// Function to draw particles on canvas
function drawParticles() {
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = '#5ed628'; // Particle color

  particles.forEach(particle => {
    ctx.beginPath();
    ctx.arc(particle.x, particle.y, particle.size, 0, Math.PI * 2);
    ctx.fill();
  });
}

// Function to update particle positions
function updateParticles() {
  particles.forEach(particle => {
    particle.x += particle.speedX;
    particle.y += particle.speedY;

    // Wrap particles around the screen edges
    if (particle.x > canvas.width) {
      particle.x = 0;
    }
    if (particle.x < 0) {
      particle.x = canvas.width;
    }
    if (particle.y > canvas.height) {
      particle.y = 0;
    }
    if (particle.y < 0) {
      particle.y = canvas.height;
    }
  });
}

// Function to animate particles
function animateParticles() {
  requestAnimationFrame(animateParticles);
  updateParticles();
  drawParticles();
}

// Generate initial particles
for (let i = 0; i < 100; i++) {
  createParticle();
}

// Start the animation
animateParticles();
