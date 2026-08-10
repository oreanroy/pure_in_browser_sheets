import { deepClone } from "./utils.js";

export class History {
  constructor(limit = 100) {
    this.limit = limit;
    this.stack = [];
    this.index = -1;
  }

  clear() {
    this.stack = [];
    this.index = -1;
  }

  push(snapshot) {
    this.stack = this.stack.slice(0, this.index + 1);
    this.stack.push(deepClone(snapshot));
    if (this.stack.length > this.limit) {
      this.stack.shift();
    } else {
      this.index++;
    }
  }

  canUndo() {
    return this.index > 0;
  }

  canRedo() {
    return this.index < this.stack.length - 1;
  }

  undo() {
    if (!this.canUndo()) return null;
    this.index--;
    return deepClone(this.stack[this.index]);
  }

  redo() {
    if (!this.canRedo()) return null;
    this.index++;
    return deepClone(this.stack[this.index]);
  }
}
