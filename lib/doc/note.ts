import {deepMerge} from '../utils/object.js';
import type {Comment, CommentEditAs, RichText} from '../../index.js';

/** Note payload: plain string or full Comment-like object. */
export type NoteInput = string | Comment | Record<string, unknown>;

export interface NoteModel {
  type: 'note';
  note: Comment | {texts: Array<{text: string}>};
}

export interface NoteDefaultConfigs {
  note: {
    margins: {
      insetmode: 'auto' | 'custom';
      inset: number[];
    };
    protection: {
      locked: 'True' | 'False';
      lockText: 'True' | 'False';
    };
    editAs: CommentEditAs;
  };
}

class Note {
  note: NoteInput | undefined;

  static DEFAULT_CONFIGS: NoteDefaultConfigs = {
    note: {
      margins: {
        insetmode: 'auto',
        inset: [0.13, 0.13, 0.25, 0.25],
      },
      protection: {
        locked: 'True',
        lockText: 'True',
      },
      editAs: 'absolute',
    },
  };

  constructor(note?: NoteInput) {
    this.note = note;
  }

  get model(): NoteModel {
    let value: NoteModel | null = null;
    switch (typeof this.note) {
      case 'string':
        value = {
          type: 'note',
          note: {
            texts: [
              {
                text: this.note,
              },
            ],
          },
        };
        break;
      default:
        value = {
          type: 'note',
          note: this.note as Comment,
        };
        break;
    }
    // Suitable for all cell comments
    return deepMerge({}, Note.DEFAULT_CONFIGS, value) as unknown as NoteModel;
  }

  set model(value: {note: Comment & {texts?: RichText[]}}) {
    const {note} = value;
    const {texts} = note;
    if (texts && texts.length === 1 && Object.keys(texts[0]).length === 1) {
      this.note = texts[0].text;
    } else {
      this.note = note;
    }
  }

  static fromModel(model: {note: Comment & {texts?: RichText[]}}): Note {
    const note = new Note();
    note.model = model;
    return note;
  }
}

export default Note;
export {Note};
