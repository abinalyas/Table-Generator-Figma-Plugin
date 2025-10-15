// Function to clean up any components created outside the generated table
export function cleanupExternalComponents() {
    try {
        // Only work with the current page to avoid loading all pages
        console.log('🧹 Cleaning up external components on current page only');
        
        // Find and remove any "Simple Divider" components that might be outside the table
        const simpleDividers = figma.currentPage.findAll(node => 
            node.type === "COMPONENT" && node.name === "Simple Divider"
        );
        
        simpleDividers.forEach(divider => {
            // Only remove if it's not inside a generated table
            let parent = divider.parent;
            let shouldRemove = true;
            
            while (parent) {
                if (parent.getPluginData('isGeneratedTable') === 'true') {
                    shouldRemove = false;
                    break;
                }
                parent = parent.parent;
            }
            
            if (shouldRemove) {
                divider.remove();
                console.log('🧹 Cleaned up external Simple Divider component');
            }
        });
        
        // Find and remove any "Data table select cell item" component sets that might be outside the table
        const checkboxComponents = figma.currentPage.findAll(node => 
            node.type === "COMPONENT_SET" && node.name === "Data table select cell item"
        );
        
        checkboxComponents.forEach(componentSet => {
            // Only remove if it's not inside a generated table
            let parent = componentSet.parent;
            let shouldRemove = true;
            
            while (parent) {
                if (parent.getPluginData('isGeneratedTable') === 'true') {
                    shouldRemove = false;
                    break;
                }
                parent = parent.parent;
            }
            
            if (shouldRemove) {
                componentSet.remove();
                console.log('🧹 Cleaned up external checkbox component set');
            }
        });
    } catch (error) {
        console.error('❌ Error cleaning up external components:', error);
    }
}