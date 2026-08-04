// =======================================================================================================
// StreamConverter
//
// convert between encoding schemes in a stream
// Work in Progress - Will complete this at some point
import {alloc, copy, from as bytesFrom} from '../utils/bytes.js';

// WIP encoding converter (never assigned a real jconv implementation yet)
const jconv: {
  convert: (data: Uint8Array, from: string, to: string) => Uint8Array;
} | null = null;

export interface StreamConverterOptions {
  innerEncoding?: string;
  outerEncoding?: string;
  innerBOM?: Uint8Array | null;
  outerBOM?: Uint8Array | null;
}

/** Minimal stream-like surface used by StreamConverter */
export interface ConvertibleStream {
  addListener(event: string, handler: (...args: unknown[]) => void): unknown;
  removeListener(event: string, handler: (...args: unknown[]) => void): unknown;
  write(
    data?: unknown,
    encoding?: string | ((...args: unknown[]) => void),
    callback?: (...args: unknown[]) => void,
  ): unknown;
  pipe(destination: unknown, options?: unknown): unknown;
  close?(): void;
  on(type: string, callback: (...args: unknown[]) => void): unknown;
  once(type: string, callback: (...args: unknown[]) => void): unknown;
  end(
    chunk?: unknown,
    encoding?: string | ((...args: unknown[]) => void),
    callback?: (...args: unknown[]) => void,
  ): unknown;
  emit(type: string, value?: unknown): unknown;
}

class StreamConverter {
  inner: ConvertibleStream;
  innerEncoding: string;
  outerEncoding: string;
  innerBOM: Uint8Array | null;
  outerBOM: Uint8Array | null;
  writeStarted: boolean;

  constructor(inner: ConvertibleStream, options?: StreamConverterOptions) {
    this.inner = inner;

    options = options || {};
    this.innerEncoding = (options.innerEncoding || 'UTF8').toUpperCase();
    this.outerEncoding = (options.outerEncoding || 'UTF8').toUpperCase();

    this.innerBOM = options.innerBOM || null;
    this.outerBOM = options.outerBOM || null;

    this.writeStarted = false;
  }

  convertInwards(data: string | Uint8Array | null | undefined): Uint8Array | null | undefined {
    if (data) {
      if (typeof data === 'string') {
        data = bytesFrom(data, this.outerEncoding);
      }

      if (this.innerEncoding !== this.outerEncoding) {
        data = jconv!.convert(data, this.outerEncoding, this.innerEncoding);
      }
    }

    return data as Uint8Array | null | undefined;
  }

  convertOutwards(data: string | Uint8Array): Uint8Array {
    if (typeof data === 'string') {
      data = bytesFrom(data, this.innerEncoding);
    }

    if (this.innerEncoding !== this.outerEncoding) {
      data = jconv!.convert(data, this.innerEncoding, this.outerEncoding);
    }
    return data;
  }

  addListener(event: string, handler: (...args: unknown[]) => void): unknown {
    return this.inner.addListener(event, handler);
  }

  removeListener(event: string, handler: (...args: unknown[]) => void): unknown {
    return this.inner.removeListener(event, handler);
  }

  write(
    data?: string | Uint8Array,
    encoding?: string | ((...args: unknown[]) => void),
    callback?: (...args: unknown[]) => void,
  ): unknown {
    if (encoding instanceof Function) {
      callback = encoding;
      encoding = undefined;
    }

    if (!this.writeStarted) {
      // if inner encoding has BOM, write it now
      if (this.innerBOM) {
        this.inner.write(this.innerBOM);
      }

      // if outer encoding has BOM, delete it now
      if (this.outerBOM && data) {
        const bytes = typeof data === 'string' ? bytesFrom(data, this.outerEncoding) : data;
        if (bytes.length <= this.outerBOM.length) {
          if (callback) {
            callback();
          }
          return;
        }
        const bomless = alloc(bytes.length - this.outerBOM.length);
        copy(bytes, bomless, 0, this.outerBOM.length, bytes.length);
        data = bomless;
      }

      this.writeStarted = true;
    }

    return this.inner.write(
      this.convertInwards(data),
      encoding ? this.innerEncoding : undefined,
      callback,
    );
  }

  read(): void {
    // TBD
  }

  pipe(destination: ConvertibleStream, options?: unknown): unknown {
    const reverseConverter = new StreamConverter(destination, {
      innerEncoding: this.outerEncoding,
      outerEncoding: this.innerEncoding,
      innerBOM: this.outerBOM,
      outerBOM: this.innerBOM,
    });

    return this.inner.pipe(reverseConverter, options);
  }

  close(): void {
    this.inner.close?.();
  }

  on(type: string, callback: (...args: unknown[]) => void): this {
    switch (type) {
      case 'data':
        this.inner.on('data', (chunk: unknown) => {
          callback(this.convertOutwards(chunk as string | Uint8Array));
        });
        return this;
      default:
        this.inner.on(type, callback);
        return this;
    }
  }

  once(type: string, callback: (...args: unknown[]) => void): unknown {
    return this.inner.once(type, callback);
  }

  end(
    chunk?: string | Uint8Array,
    encoding?: string | ((...args: unknown[]) => void),
    callback?: (...args: unknown[]) => void,
  ): unknown {
    return this.inner.end(this.convertInwards(chunk), this.innerEncoding, callback);
  }

  emit(type: string, value?: unknown): unknown {
    return this.inner.emit(type, value);
  }
}

export default StreamConverter;
export {StreamConverter};
