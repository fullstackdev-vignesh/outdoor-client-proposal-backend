const fs = require('fs');
const path = require('path');
const { getRouteMapBuffer } = require(path.join(process.cwd(), 'src', 'utils', 'mapService.js'));

async function testPublicis() {
  const res = await getRouteMapBuffer({
    fromLat: 13.030,
    fromLng: 80.060,
    toLat: 13.050,
    toLng: 80.280,
    width: 640,
    height: 480
  });

  if (res && res.buffer) {
    fs.writeFileSync('.artifacts/publicis_route_test.png', res.buffer);
    console.log('Publicis map generated successfully!');
    console.log('Distance:', res.distanceText, 'Duration:', res.durationText);
  } else {
    console.error('Failed to generate map');
  }
}

testPublicis();
