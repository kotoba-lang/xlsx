/**
 * Comments — cell-level notes and threaded comments.
 *
 * Provides a {@link CommentStore} that manages comments keyed by cell
 * reference. Each comment supports threaded replies.
 */
import type { CellRef } from "./ooxml-parser";

/** A top-level comment attached to a cell. */
export interface CellComment {
  /** The cell reference this comment is attached to (e.g. "A1"). */
  ref: CellRef;
  /** Display name of the comment author. */
  author: string;
  /** Comment body text. */
  text: string;
  /** ISO 8601 timestamp of creation. */
  createdAt: string;
  /** Threaded replies to this comment, in chronological order. */
  replies: CellCommentReply[];
}

/** A reply within a comment thread. */
export interface CellCommentReply {
  /** Display name of the reply author. */
  author: string;
  /** Reply body text. */
  text: string;
  /** ISO 8601 timestamp of creation. */
  createdAt: string;
}

/**
 * Comment store — manages cell comments as a Map keyed by {@link CellRef}.
 *
 * Supports add, get, delete, reply, and JSON serialisation for persistence.
 */
export class CommentStore {
  private comments = new Map<CellRef, CellComment>();

  /**
   * Add a new comment to a cell. Overwrites any existing comment on the
   * same cell reference.
   *
   * @param ref - The cell reference.
   * @param author - The comment author name.
   * @param text - The comment body.
   * @returns The newly created comment.
   */
  add(ref: CellRef, author: string, text: string): CellComment {
    const comment: CellComment = {
      ref,
      author,
      text,
      createdAt: new Date().toISOString(),
      replies: [],
    };
    this.comments.set(ref, comment);
    return comment;
  }

  /**
   * Retrieve the comment for a given cell, if any.
   *
   * @param ref - The cell reference.
   * @returns The comment or `undefined`.
   */
  get(ref: CellRef): CellComment | undefined {
    return this.comments.get(ref);
  }

  /**
   * Delete the comment on a cell.
   *
   * @param ref - The cell reference.
   * @returns `true` if a comment was removed, `false` if none existed.
   */
  delete(ref: CellRef): boolean {
    return this.comments.delete(ref);
  }

  /**
   * Append a reply to an existing comment thread.
   *
   * If no comment exists at `ref`, this is a no-op (caller should add a
   * comment first).
   *
   * @param ref - The cell reference whose comment receives the reply.
   * @param author - The reply author name.
   * @param text - The reply body.
   */
  reply(ref: CellRef, author: string, text: string): void {
    const comment = this.comments.get(ref);
    if (!comment) return;
    comment.replies.push({
      author,
      text,
      createdAt: new Date().toISOString(),
    });
  }

  /**
   * Return all comments as an array, ordered by cell reference.
   *
   * @returns Array of all {@link CellComment} instances.
   */
  getAll(): CellComment[] {
    return Array.from(this.comments.values()).sort((a, b) => a.ref.localeCompare(b.ref));
  }

  /**
   * Check whether a cell has an attached comment.
   *
   * @param ref - The cell reference.
   * @returns `true` if a comment exists on the cell.
   */
  hasComment(ref: CellRef): boolean {
    return this.comments.has(ref);
  }

  /**
   * Serialise the comment store to a plain JSON-compatible object.
   *
   * @returns An object with a `comments` array suitable for `JSON.stringify`.
   */
  toJSON(): object {
    return {
      comments: Array.from(this.comments.values()),
    };
  }

  /**
   * Deserialise a comment store from the output of {@link toJSON}.
   *
   * @param data - The plain object (typically parsed from JSON).
   * @returns A new {@link CommentStore} populated with the data.
   */
  static fromJSON(data: { comments?: CellComment[] }): CommentStore {
    const store = new CommentStore();
    if (data && Array.isArray(data.comments)) {
      for (const c of data.comments) {
        const comment: CellComment = {
          ref: c.ref,
          author: c.author,
          text: c.text,
          createdAt: c.createdAt,
          replies: Array.isArray(c.replies)
            ? c.replies.map((r) => ({
                author: r.author,
                text: r.text,
                createdAt: r.createdAt,
              }))
            : [],
        };
        store.comments.set(comment.ref, comment);
      }
    }
    return store;
  }
}
