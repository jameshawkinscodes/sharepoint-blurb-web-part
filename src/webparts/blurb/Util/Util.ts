// Utility file: src/utils/Util.ts

interface SortableItem {
  SortWeight: number;
}

export class Util {
  /**
   * Generates a unique ID.
   * @returns A unique string ID.
   */
  public static GenerateId(): string {
    return 'id-' + Math.random().toString(36).substr(2, 16);
  }

  /**
   * Calculates the new sort weight for an item when reordered.
   * @param items - The array of sortable items.
   * @param newIndex - The new index of the item being reordered.
   * @param oldIndex - The original index of the item (optional).
   * @returns The new sort weight for the item.
   */
  public static CalculateNewSortWeight(
    items: SortableItem[],
    newIndex: number,
    oldIndex?: number
  ): number {
    // Sort items by their SortWeight property
    const sortedItems = [...items].sort((a, b) => a.SortWeight - b.SortWeight);

    let newSortWeight = 0;

    if (newIndex === 0) {
      // If the new index is at the start, calculate a sort weight less than the first item
      newSortWeight = sortedItems[0].SortWeight - 1;
    } else if (newIndex === sortedItems.length) {
      // If the new index is at the end, calculate a sort weight greater than the last item
      newSortWeight = sortedItems[sortedItems.length - 1].SortWeight + 1;
    } else {
      // Otherwise, calculate a sort weight between the surrounding items
      const prevItem = sortedItems[newIndex - 1];
      const nextItem = sortedItems[newIndex];
      newSortWeight = (prevItem.SortWeight + nextItem.SortWeight) / 2;
    }

    return newSortWeight;
  }
}
