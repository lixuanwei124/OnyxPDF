'use strict';

const crypto = require('crypto');
const PDFLib = require('pdf-lib');

const PASSWORD_PAD = Uint8Array.from([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41,
  0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80,
  0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a
]);

function md5(...items) {
  const hash = crypto.createHash('md5');
  for (const item of items) hash.update(item);
  return hash.digest();
}

function rc4(key, input) {
  const state = Uint8Array.from({ length: 256 }, (_, index) => index);
  let swap = 0;
  for (let index = 0; index < 256; index++) {
    swap = (swap + state[index] + key[index % key.length]) & 255;
    [state[index], state[swap]] = [state[swap], state[index]];
  }
  swap = 0;
  const output = new Uint8Array(input.length);
  let left = 0;
  for (let index = 0; index < input.length; index++) {
    left = (left + 1) & 255;
    swap = (swap + state[left]) & 255;
    [state[left], state[swap]] = [state[swap], state[left]];
    output[index] = input[index] ^ state[(state[left] + state[swap]) & 255];
  }
  return output;
}

function objectKey(globalKey, objectNumber, generation, aes) {
  const number = Buffer.alloc(3);
  number.writeUIntLE(Number(objectNumber) >>> 0, 0, 3);
  const generationBuffer = Buffer.alloc(2);
  generationBuffer.writeUIntLE(Number(generation) >>> 0, 0, 2);
  const extension = aes ? Buffer.from([0x73, 0x41, 0x6c, 0x54]) : Buffer.alloc(0);
  return md5(globalKey, number, generationBuffer, extension);
}

function decryptAesCbc(key, input) {
  if (input.length < 32 || input.length % 16) return null;
  const iv = input.subarray(0, 16);
  const data = input.subarray(16);
  try {
    const decipher = crypto.createDecipheriv('aes-128-cbc', key, iv);
    decipher.setAutoPadding(false);
    const plain = Buffer.concat([decipher.update(data), decipher.final()]);
    const padding = plain[plain.length - 1];
    if (padding < 1 || padding > 16) return null;
    for (let index = plain.length - padding; index < plain.length; index++) {
      if (plain[index] !== padding) return null;
    }
    return plain.subarray(0, plain.length - padding);
  } catch (_) {
    return null;
  }
}

function nameValue(value) {
  return String(value?.asString?.() ?? value?.decodeText?.() ?? value ?? '')
    .replace(/^\//, '');
}

function cryptMethods(document, encrypt) {
  const version = Math.floor(encrypt.lookupMaybe(PDFLib.PDFName.of('V'), PDFLib.PDFNumber)?.asNumber() || 0);
  if (version <= 2) return { stream: 'rc4', string: 'rc4' };
  if (version >= 5) return null;
  const filters = encrypt.lookupMaybe(PDFLib.PDFName.of('CF'), PDFLib.PDFDict);
  const method = dict => {
    if (!dict) return null;
    return /AESV2|AESV3/i.test(nameValue(dict.get(PDFLib.PDFName.of('CFM')))) ? 'aes' : 'rc4';
  };
  const streamName = nameValue(encrypt.get(PDFLib.PDFName.of('StmF'))) || 'Identity';
  const stringName = nameValue(encrypt.get(PDFLib.PDFName.of('StrF'))) || 'Identity';
  const streamFilter = filters?.lookupMaybe(PDFLib.PDFName.of(streamName), PDFLib.PDFDict);
  const stringFilter = filters?.lookupMaybe(PDFLib.PDFName.of(stringName), PDFLib.PDFDict);
  return {
    stream: streamName === 'Identity' ? 'rc4' : method(streamFilter),
    string: stringName === 'Identity' ? 'rc4' : method(stringFilter)
  };
}

function encryptionKey(document, encrypt) {
  const revision = Math.floor(encrypt.lookupMaybe(PDFLib.PDFName.of('R'), PDFLib.PDFNumber)?.asNumber() || 0);
  const version = Math.floor(encrypt.lookupMaybe(PDFLib.PDFName.of('V'), PDFLib.PDFNumber)?.asNumber() || 0);
  const lengthBits = Math.floor(encrypt.lookupMaybe(PDFLib.PDFName.of('Length'), PDFLib.PDFNumber)?.asNumber() || (version >= 2 ? 128 : 40));
  let keyLength = Math.min(16, Math.max(5, Math.ceil(lengthBits / 8) + 5));
  if (revision === 2) keyLength = 5;
  const permissions = Math.floor(encrypt.lookupMaybe(PDFLib.PDFName.of('P'), PDFLib.PDFNumber)?.asNumber() || 0);
  const permissionsBuffer = Buffer.alloc(4);
  permissionsBuffer.writeInt32LE(permissions | 0, 0);
  const ownerEntry = encrypt.lookup(PDFLib.PDFName.of('O'))?.asBytes?.() || Buffer.alloc(32);
  const trailerId = document.context.trailerInfo?.ID;
  const identifier = trailerId instanceof PDFLib.PDFArray
    ? trailerId.get(0)?.asBytes?.() || Buffer.alloc(0)
    : Array.isArray(trailerId)
      ? trailerId[0]?.asBytes?.() || Buffer.alloc(0)
      : Buffer.alloc(0);
  const encryptMetadata = encrypt.lookupMaybe(PDFLib.PDFName.of('EncryptMetadata'), PDFLib.PDFBool);
  const extra = revision >= 4 && encryptMetadata && !encryptMetadata.asBoolean()
    ? Buffer.from([0xff, 0xff, 0xff, 0xff])
    : Buffer.alloc(0);

  let key = md5(PASSWORD_PAD, Buffer.from(ownerEntry), permissionsBuffer, identifier, extra);
  if (revision >= 3) {
    for (let index = 0; index < 50; index++) {
      key = md5(key.subarray(0, keyLength));
    }
  }
  return key.subarray(0, keyLength);
}

function decryptValue(value, key, method, objectNumber, generation) {
  if (!(value instanceof PDFLib.PDFString || value instanceof PDFLib.PDFHexString)) return false;
  const encrypted = value.asBytes();
  const itemKey = objectKey(key, objectNumber, generation, method === 'aes');
  const plain = method === 'aes' ? decryptAesCbc(itemKey, encrypted) : rc4(itemKey, encrypted);
  if (!plain) return false;
  value.value = Buffer.from(plain).toString('latin1');
  Object.setPrototypeOf(value, PDFLib.PDFString.prototype);
  return true;
}

function walkStrings(value, key, method, objectNumber, generation, seen = new Set()) {
  if (!value || typeof value !== 'object' || seen.has(value)) return;
  seen.add(value);
  if (value instanceof PDFLib.PDFString || value instanceof PDFLib.PDFHexString) {
    decryptValue(value, key, method, objectNumber, generation);
    return;
  }
  if (value instanceof PDFLib.PDFDict) {
    for (const [name, child] of value.entries()) walkStrings(child, key, method, objectNumber, generation, seen);
  } else if (value instanceof PDFLib.PDFArray) {
    value.asArray().forEach((child, index) => {
      if (child instanceof PDFLib.PDFRef) return;
      walkStrings(child, key, method, objectNumber, generation, seen);
    });
  }
}

async function decryptPdf(bytes) {
  const source = bytes instanceof Uint8Array ? Buffer.from(bytes) : Buffer.from(await Promise.resolve(bytes));
  const document = await PDFLib.PDFDocument.load(source, {
    ignoreEncryption: true,
    throwOnInvalidObject: false,
    updateMetadata: false
  });
  const encryptReference = document.context.trailerInfo?.Encrypt;
  if (!encryptReference) return { ok: true, encrypted: false, bytes: new Uint8Array(source) };

  const encrypt = document.context.lookup(encryptReference);
  const filter = nameValue(encrypt?.get(PDFLib.PDFName.of('Filter')));
  if (!/Standard/i.test(filter)) return { ok: false, error: '暂不支持该 PDF 加密过滤器' };
  const revision = Math.floor(encrypt.lookupMaybe(PDFLib.PDFName.of('R'), PDFLib.PDFNumber)?.asNumber() || 0);
  if (revision >= 5) return { ok: false, error: '暂不支持 AES-256 加密 PDF 的内容编辑' };
  const methods = cryptMethods(document, encrypt);
  if (!methods?.stream || !methods?.string) return { ok: false, error: '暂不支持该 PDF 加密算法' };
  const globalKey = encryptionKey(document, encrypt);

  for (const [reference, object] of document.context.enumerateIndirectObjects()) {
    if (object instanceof PDFLib.PDFRawStream) {
      const itemKey = objectKey(globalKey, reference.objectNumber, reference.generation, methods.stream === 'aes');
      const encrypted = object.contents;
      const plain = methods.stream === 'aes' ? decryptAesCbc(itemKey, encrypted) : rc4(itemKey, encrypted);
      if (!plain) return { ok: false, error: 'PDF 解密失败：内容流校验不通过' };
      object.contents = new Uint8Array(plain);
      walkStrings(object.dict, globalKey, methods.string, reference.objectNumber, reference.generation);
    } else {
      walkStrings(object, globalKey, methods.string, reference.objectNumber, reference.generation);
    }
  }
  document.context.trailerInfo.Encrypt = undefined;
  const saved = await document.save({ useObjectStreams: false, updateFieldAppearances: false });
  return { ok: true, encrypted: true, bytes: new Uint8Array(saved) };
}

module.exports = { decryptPdf };
