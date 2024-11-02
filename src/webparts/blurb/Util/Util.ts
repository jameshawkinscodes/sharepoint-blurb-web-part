// Utility file: src/utils/Util.ts

export class Util {
    // Function to generate a unique ID
    public static GenerateId(): string {
      return 'id-' + Math.random().toString(36).substr(2, 16);
    }
  
    // Function to calculate the sort weight for blurbs when they are reordered
    public static CalculateNewSortWeight(items: any[], newIndex: number, oldIndex?: number): number {
      const sortedItems = [...items].sort((a, b) => a.SortWeight - b.SortWeight);
  
      let newSortWeight = 0;
  
      if (newIndex === 0) {
        // If the new index is at the start, calculate a sort weight less than the first item
        newSortWeight = sortedItems[0].SortWeight - 1;
      } else if (newIndex === sortedItems.length - 1) {
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
  