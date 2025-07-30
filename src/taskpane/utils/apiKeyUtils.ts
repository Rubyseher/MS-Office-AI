/**
 * Utility functions for managing API keys in Office Add-ins
 */

/**
 * Retrieves the stored API key from Office roaming settings or localStorage
 */
export const getStoredApiKey = async (): Promise<string> => {
  try {
    // Try Office roaming settings first (syncs across devices)
    if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
      const apiKey = Office.context.roamingSettings.get("gemini_api_key");
      return apiKey || "";
    }
    
    // Fallback to localStorage for development
    return localStorage.getItem("gemini_api_key") || "";
  } catch (error) {
    console.error("Error retrieving API key:", error);
    return "";
  }
};

/**
 * Stores the API key to Office roaming settings or localStorage
 */
export const storeApiKey = async (apiKey: string): Promise<void> => {
  try {
    // Try Office roaming settings first (syncs across devices)
    if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
      Office.context.roamingSettings.set("gemini_api_key", apiKey);
      await Office.context.roamingSettings.saveAsync();
    } else {
      // Fallback to localStorage for development
      localStorage.setItem("gemini_api_key", apiKey);
    }
  } catch (error) {
    console.error("Error storing API key:", error);
    throw error;
  }
};

/**
 * Removes the stored API key
 */
export const removeStoredApiKey = async (): Promise<void> => {
  try {
    // Try Office roaming settings first
    if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
      Office.context.roamingSettings.remove("gemini_api_key");
      await Office.context.roamingSettings.saveAsync();
    } else {
      // Fallback to localStorage for development
      localStorage.removeItem("gemini_api_key");
    }
  } catch (error) {
    console.error("Error removing API key:", error);
    throw error;
  }
};

/**
 * Validates if an API key appears to be in the correct format
 */
export const validateApiKey = (apiKey: string): boolean => {
  // More lenient validation for Gemini API key format
  // Allow various formats as long as it's a reasonable length and contains valid characters
  if (!apiKey || typeof apiKey !== 'string') return false;
  
  // Must be at least 15 characters and contain alphanumeric/underscore/dash
  return apiKey.length >= 15 && /^[A-Za-z0-9_-]+$/.test(apiKey);
};
