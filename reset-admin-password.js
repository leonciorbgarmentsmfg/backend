import { randomBytes, scryptSync } from 'node:crypto';
import db from './db/init.js';

function readHidden(prompt) {
  return new Promise((resolve, reject) => {
    const input = process.stdin;
    if (!input.isTTY || typeof input.setRawMode !== 'function') {
      reject(new Error('Run this command in an interactive terminal.'));
      return;
    }

    let value = '';
    let finished = false;
    process.stdout.write(prompt);
    input.setRawMode(true);
    input.resume();

    function finish(error) {
      if (finished) return;
      finished = true;
      input.removeListener('data', handleInput);
      input.setRawMode(false);
      input.pause();
      process.stdout.write('\n');
      if (error) reject(error);
      else resolve(value);
    }

    function handleInput(chunk) {
      for (const character of chunk.toString('utf8')) {
        if (character === '\u0003') {
          finish(new Error('Password reset cancelled.'));
          return;
        }
        if (character === '\r' || character === '\n') {
          finish();
          return;
        }
        if (character === '\u007f' || character === '\b') {
          if (value.length > 0) {
            value = Array.from(value).slice(0, -1).join('');
            process.stdout.write('\b \b');
          }
          continue;
        }
        if (character >= ' ') {
          value += character;
          process.stdout.write('*');
        }
      }
    }

    input.on('data', handleInput);
  });
}

try {
  const password = await readHidden('New Administrator password (8+ characters): ');
  if (password.length < 8) {
    throw new Error('Password must be at least 8 characters.');
  }
  const confirmation = await readHidden('Confirm new password: ');
  if (password !== confirmation) {
    throw new Error('Passwords do not match.');
  }

  const admin = db
    .prepare('SELECT id FROM users WHERE is_protected = 1 AND role = ?')
    .get('admin');
  if (!admin) {
    throw new Error('The protected Administrator account was not found.');
  }

  const salt = randomBytes(16).toString('hex');
  const passwordHash = `${salt}:${scryptSync(password, salt, 64).toString('hex')}`;
  const resetPassword = db.transaction(() => {
    db.prepare(
      'UPDATE users SET password_hash = ?, must_change_password = 1 WHERE id = ?',
    ).run(passwordHash, admin.id);
    db.prepare('DELETE FROM sessions WHERE user_id = ?').run(admin.id);
  });
  resetPassword.immediate();
  process.stdout.write('Administrator password reset. Sign in as Administrator and set a new password.\n');
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
} finally {
  db.close();
}
