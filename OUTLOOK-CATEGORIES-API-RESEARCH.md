# Outlook Categories API Research Summary

## Overview
The Office.js API provides comprehensive support for managing categories (tags/labels) in Outlook add-ins. Categories allow users to organize messages and appointments with color-coded tags. This document summarizes the available methods, code examples, and limitations.

---

## 1. Setting Categories/Tags on Emails Programmatically

### API Location
Categories are managed through two main interfaces:
- **`Office.context.mailbox.masterCategories`** - Manages the master list of categories
- **`Office.context.mailbox.item.categories`** - Manages categories applied to individual items

### Adding Categories to an Email

```javascript
// Categories must exist in master list first
const categoriesToAdd = ["Urgent!", "Important"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### Getting Categories from an Email

```javascript
Office.context.mailbox.item.categories.getAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const categories = asyncResult.value;
        // Note: Returns null or empty array if no categories exist
        if (categories && categories.length > 0) {
            console.log("Categories:");
            categories.forEach(function(item) {
                console.log("-- " + JSON.stringify(item));
            });
        } else {
            console.log("There are no categories assigned to this item.");
        }
    }
});
```

### Removing Categories from an Email

```javascript
const categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

---

## 2. Creating Custom Categories with Custom Colors

### Creating Master Categories

Before applying categories to items, they must exist in the master list. You can create custom categories with specific colors:

```javascript
const masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0  // Red
    },
    {
        "displayName": "Important",
        "color": Office.MailboxEnums.CategoryColor.Preset1  // Orange
    },
    {
        "displayName": "Follow Up",
        "color": Office.MailboxEnums.CategoryColor.Preset4  // Green
    }
];

Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### Available Color Presets

The `Office.MailboxEnums.CategoryColor` enum provides 25 preset colors:

- `None` - Default color or no color mapped
- `Preset0` - Red
- `Preset1` - Orange
- `Preset2` - Brown
- `Preset3` - Yellow
- `Preset4` - Green
- `Preset5` - Teal
- `Preset6` - Olive
- `Preset7` - Blue
- `Preset8` - Purple
- `Preset9` - Cranberry
- `Preset10` - Steel
- `Preset11` - DarkSteel
- `Preset12` - Gray
- `Preset13` - DarkGray
- `Preset14` - Black
- `Preset15` - DarkRed
- `Preset16` - DarkOrange
- `Preset17` - DarkBrown
- `Preset18` - DarkYellow
- `Preset19` - DarkGreen
- `Preset20` - DarkTeal
- `Preset21` - DarkOlive
- `Preset22` - DarkBlue
- `Preset23` - DarkPurple
- `Preset24` - DarkCranberry

**Important Notes:**
- Multiple categories can share the same color
- Each category name must be unique
- The actual rendered color may vary slightly by Outlook client
- Earlier versions of Outlook on Mac (before 16.78) had incorrect preset colors

### Getting Master Categories

```javascript
Office.context.mailbox.masterCategories.getAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function(item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### Removing Master Categories

```javascript
const masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

---

## 3. Modifying Categories on Existing Emails

### Complete Workflow Example

```javascript
// Step 1: Ensure category exists in master list
function ensureCategoryExists(categoryName, color, callback) {
    Office.context.mailbox.masterCategories.getAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const masterCategories = asyncResult.value;
            const exists = masterCategories.some(cat => cat.displayName === categoryName);
            
            if (!exists) {
                // Create the category
                const newCategory = [{
                    displayName: categoryName,
                    color: color
                }];
                
                Office.context.mailbox.masterCategories.addAsync(newCategory, function(addResult) {
                    if (addResult.status === Office.AsyncResultStatus.Succeeded) {
                        callback(null, true);
                    } else {
                        callback(addResult.error, false);
                    }
                });
            } else {
                callback(null, true);
            }
        } else {
            callback(asyncResult.error, false);
        }
    });
}

// Step 2: Get current categories on item
function getCurrentCategories(callback) {
    Office.context.mailbox.item.categories.getAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const categories = asyncResult.value || [];
            callback(null, categories.map(cat => cat.displayName));
        } else {
            callback(asyncResult.error, null);
        }
    });
}

// Step 3: Add or update categories
function updateItemCategories(categoryName, color, callback) {
    ensureCategoryExists(categoryName, color, function(err, success) {
        if (err) {
            callback(err);
            return;
        }
        
        getCurrentCategories(function(err, currentCategories) {
            if (err) {
                callback(err);
                return;
            }
            
            // Only add if not already present
            if (!currentCategories.includes(categoryName)) {
                Office.context.mailbox.item.categories.addAsync([categoryName], function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        callback(null, "Category added successfully");
                    } else {
                        callback(asyncResult.error);
                    }
                });
            } else {
                callback(null, "Category already exists on item");
            }
        });
    });
}

// Usage
updateItemCategories("Urgent!", Office.MailboxEnums.CategoryColor.Preset0, function(err, result) {
    if (err) {
        console.error("Error:", err);
    } else {
        console.log(result);
    }
});
```

### Replacing All Categories

```javascript
// Remove all existing categories, then add new ones
function replaceAllCategories(newCategoryNames, callback) {
    // First, get current categories
    Office.context.mailbox.item.categories.getAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const currentCategories = asyncResult.value || [];
            const currentNames = currentCategories.map(cat => cat.displayName);
            
            // Remove all existing
            if (currentNames.length > 0) {
                Office.context.mailbox.item.categories.removeAsync(currentNames, function(removeResult) {
                    if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
                        // Add new categories
                        Office.context.mailbox.item.categories.addAsync(newCategoryNames, callback);
                    } else {
                        callback(removeResult.error);
                    }
                });
            } else {
                // No existing categories, just add new ones
                Office.context.mailbox.item.categories.addAsync(newCategoryNames, callback);
            }
        } else {
            callback(asyncResult.error);
        }
    });
}
```

---

## 4. Requirements and Limitations

### Requirement Set
- **Requirement Set 1.8** - Categories API was introduced in this requirement set
- Check client/platform support: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets

### Permission Requirements

#### For Managing Master Categories:
- **Required Permission:** `ReadWriteMailbox` (read/write mailbox)
- **Manifest Configuration:**
  - **Add-in only manifest:** Set `<Permissions>` element to `ReadWriteMailbox`
  - **Unified manifest for Microsoft 365:** Set `"name"` property in `"authorization.permissions.resourceSpecific"` array to `"Mailbox.ReadWrite.User"`

#### For Managing Item Categories:
- **Minimum Permission:** `ReadItem` (to get categories)
- **Required Permission:** `ReadWriteItem` (to add/remove categories)

### Limitations

1. **Compose Mode Restrictions:**
   - **Outlook on the web:** Cannot manage categories applied to messages in Compose mode
   - **New Outlook on Windows:** Cannot manage categories applied to messages in Compose mode
   - **Appointments:** Categories can be managed in Compose mode
   - **Read Mode:** Full category management is available

2. **Delegate/Shared Mailbox Scenarios:**
   - Delegates can **view** the master category list
   - Delegates **cannot** add or remove categories from the master list
   - Delegates can manage categories on individual items

3. **Category Name Uniqueness:**
   - Each category name must be unique in the master list
   - Multiple categories can share the same color
   - Attempting to add a duplicate category name will result in `DuplicateCategory` error

4. **Error Handling:**
   - `InvalidCategory`: Invalid categories were provided when adding to an item
   - `DuplicateCategory`: Category already exists in master list
   - `PermissionDenied`: User doesn't have permission to perform the action

5. **Return Value Handling:**
   - `getAsync()` may return `null` or an empty array when no categories exist
   - Always check for both cases: `if (categories && categories.length > 0)`

6. **Color Rendering:**
   - Actual colors depend on Outlook client rendering
   - Colors are consistent across Outlook on web, Windows (new and classic), and Mac (16.78+)
   - Earlier Mac versions (before 16.78) had incorrect preset colors

---

## 5. Complete Example: Full Category Management

```javascript
// Complete category management utility
const CategoryManager = {
    // Ensure category exists in master list, create if needed
    ensureMasterCategory: function(categoryName, color, callback) {
        Office.context.mailbox.masterCategories.getAsync(function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const masterCategories = asyncResult.value || [];
                const exists = masterCategories.some(cat => cat.displayName === categoryName);
                
                if (!exists) {
                    const newCategory = [{
                        displayName: categoryName,
                        color: color || Office.MailboxEnums.CategoryColor.Preset0
                    }];
                    
                    Office.context.mailbox.masterCategories.addAsync(newCategory, function(addResult) {
                        callback(addResult.status === Office.AsyncResultStatus.Succeeded ? null : addResult.error);
                    });
                } else {
                    callback(null); // Already exists
                }
            } else {
                callback(asyncResult.error);
            }
        });
    },
    
    // Get all master categories
    getMasterCategories: function(callback) {
        Office.context.mailbox.masterCategories.getAsync(function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                callback(null, asyncResult.value || []);
            } else {
                callback(asyncResult.error, null);
            }
        });
    },
    
    // Get categories on current item
    getItemCategories: function(callback) {
        Office.context.mailbox.item.categories.getAsync(function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                callback(null, asyncResult.value || []);
            } else {
                callback(asyncResult.error, null);
            }
        });
    },
    
    // Add category to current item (ensures master category exists first)
    addCategoryToItem: function(categoryName, color, callback) {
        const self = this;
        self.ensureMasterCategory(categoryName, color, function(err) {
            if (err) {
                callback(err);
                return;
            }
            
            Office.context.mailbox.item.categories.addAsync([categoryName], function(asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    callback(null, "Category added successfully");
                } else {
                    callback(asyncResult.error);
                }
            });
        });
    },
    
    // Remove category from current item
    removeCategoryFromItem: function(categoryName, callback) {
        Office.context.mailbox.item.categories.removeAsync([categoryName], function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                callback(null, "Category removed successfully");
            } else {
                callback(asyncResult.error);
            }
        });
    },
    
    // Set specific categories on item (replaces all existing)
    setItemCategories: function(categoryNames, colors, callback) {
        const self = this;
        let categoriesCreated = 0;
        const totalCategories = categoryNames.length;
        
        if (totalCategories === 0) {
            // Remove all categories
            self.getItemCategories(function(err, currentCategories) {
                if (err) {
                    callback(err);
                    return;
                }
                
                const currentNames = (currentCategories || []).map(cat => cat.displayName);
                if (currentNames.length === 0) {
                    callback(null, "No categories to remove");
                    return;
                }
                
                Office.context.mailbox.item.categories.removeAsync(currentNames, callback);
            });
            return;
        }
        
        // Ensure all master categories exist
        categoryNames.forEach(function(name, index) {
            const color = colors && colors[index] ? colors[index] : Office.MailboxEnums.CategoryColor.Preset0;
            self.ensureMasterCategory(name, color, function(err) {
                categoriesCreated++;
                if (categoriesCreated === totalCategories) {
                    // All categories ensured, now set them on item
                    self.getItemCategories(function(err, currentCategories) {
                        if (err) {
                            callback(err);
                            return;
                        }
                        
                        const currentNames = (currentCategories || []).map(cat => cat.displayName);
                        const toRemove = currentNames.filter(name => !categoryNames.includes(name));
                        const toAdd = categoryNames.filter(name => !currentNames.includes(name));
                        
                        function addCategories() {
                            if (toAdd.length > 0) {
                                Office.context.mailbox.item.categories.addAsync(toAdd, function(addResult) {
                                    callback(addResult.status === Office.AsyncResultStatus.Succeeded ? null : addResult.error);
                                });
                            } else {
                                callback(null, "Categories already set");
                            }
                        }
                        
                        if (toRemove.length > 0) {
                            Office.context.mailbox.item.categories.removeAsync(toRemove, function(removeResult) {
                                if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
                                    addCategories();
                                } else {
                                    callback(removeResult.error);
                                }
                            });
                        } else {
                            addCategories();
                        }
                    });
                }
            });
        });
    }
};

// Usage Examples:
// CategoryManager.addCategoryToItem("Urgent!", Office.MailboxEnums.CategoryColor.Preset0, function(err, result) {
//     if (!err) console.log(result);
// });

// CategoryManager.setItemCategories(
//     ["Urgent!", "Important", "Follow Up"],
//     [Office.MailboxEnums.CategoryColor.Preset0, Office.MailboxEnums.CategoryColor.Preset1, Office.MailboxEnums.CategoryColor.Preset4],
//     function(err, result) {
//         if (!err) console.log(result);
//     }
// );
```

---

## 6. Additional Resources

- **Official Documentation:** https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/categories
- **API Reference - Categories:** https://learn.microsoft.com/en-us/javascript/api/outlook/office.categories
- **API Reference - MasterCategories:** https://learn.microsoft.com/en-us/javascript/api/outlook/office.mastercategories
- **API Reference - CategoryColor:** https://learn.microsoft.com/en-us/javascript/api/outlook/office.mailboxenums.categorycolor
- **Requirement Sets:** https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets
- **Permissions Guide:** https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions
- **Script Lab Samples:** Install Script Lab for Outlook add-in to try interactive samples

---

## Summary

The Office.js Categories API provides robust functionality for managing categories in Outlook add-ins:

✅ **Can Do:**
- Create custom categories with specific colors
- Apply multiple categories to emails/appointments
- Read, add, and remove categories programmatically
- Manage master category list
- Works in Read mode for all item types
- Works in Compose mode for appointments

❌ **Cannot Do:**
- Manage categories on messages in Compose mode (web/new Windows)
- Create categories with custom RGB colors (only preset colors available)
- Manage master categories as a delegate
- Use categories API without requirement set 1.8

The API is well-documented and provides comprehensive error handling for edge cases. Always ensure categories exist in the master list before applying them to items.
