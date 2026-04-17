import fs from 'node:fs';
import path from 'node:path';

const fromDir = path.join(process.cwd(), 'src', 'renderer');
const toDir = path.join(process.cwd(), 'build', 'renderer');

fs.mkdirSync(toDir, { recursive: true });
for (const file of ['index.html', 'styles.css']) {
  fs.copyFileSync(path.join(fromDir, file), path.join(toDir, file));
}
