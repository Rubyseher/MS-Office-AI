import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Input,
  Field,
  Spinner,
  Toast,
  ToastBody,
  ToastTitle,
  useToastController,
  useId,
} from "@fluentui/react-components";
import { Settings24Regular, Eye24Regular, EyeOff24Regular } from "@fluentui/react-icons";
import * as React from "react";
import { storeApiKey, validateApiKey } from "../utils/apiKeyUtils";

interface ApiKeyManagerProps {
  onApiKeyChange: (apiKey: string) => void;
  currentApiKey: string;
}

const ApiKeyManager: React.FC<ApiKeyManagerProps> = ({ onApiKeyChange, currentApiKey }) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [apiKey, setApiKey] = React.useState(currentApiKey);
  const [showApiKey, setShowApiKey] = React.useState(false);
  const [isSaving, setIsSaving] = React.useState(false);
  const toasterId = useId("apikey-toaster");
  const { dispatchToast } = useToastController(toasterId);

  const showToast = (
    title: string,
    body: string,
    intent: "info" | "success" | "warning" | "error"
  ) => {
    dispatchToast(
      <Toast>
        <ToastTitle>{title}</ToastTitle>
        <ToastBody>{body}</ToastBody>
      </Toast>,
      { intent }
    );
  };

  const handleSave = async () => {
    console.log("Attempting to save API key:", apiKey?.length, "characters");
    
    if (!apiKey.trim()) {
      showToast("Error", "Please enter a valid API key", "error");
      return;
    }

    const isValid = validateApiKey(apiKey);
    console.log("API key validation result:", isValid);
    
    if (!isValid) {
      showToast("Error", `Please enter a valid Gemini API key (at least 15 characters, alphanumeric)`, "error");
      return;
    }

    setIsSaving(true);
    try {
      await storeApiKey(apiKey);
      onApiKeyChange(apiKey);
      setIsOpen(false);
      showToast("Success", "API key saved successfully", "success");
      console.log("API key saved successfully");
    } catch (error) {
      console.error("Error saving API key:", error);
      showToast("Error", "Failed to save API key", "error");
    } finally {
      setIsSaving(false);
    }
  };

  const handleClose = () => {
    setApiKey(currentApiKey); // Reset to current value if cancelled
    setIsOpen(false);
  };

  React.useEffect(() => {
    setApiKey(currentApiKey);
  }, [currentApiKey]);

  const maskedApiKey = apiKey ? `${"*".repeat(Math.max(0, apiKey.length - 8))}${apiKey.slice(-8)}` : "";

  return (
    <>
      <Dialog open={isOpen} onOpenChange={(_, data) => setIsOpen(data.open)}>
        <DialogTrigger disableButtonEnhancement>
          <Button
            icon={<Settings24Regular />}
            appearance="subtle"
            size="small"
            title="API Key Settings"
          />
        </DialogTrigger>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>API Key Settings</DialogTitle>
            <DialogContent>
              <div className="space-y-4">
                <div>
                  <p className="text-sm text-gray-600 dark:text-gray-300 mb-4">
                    Enter your Google Gemini API key to use AI features. You can get a free API key from{" "}
                    <a
                      href="https://aistudio.google.com/app/apikey"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="text-blue-600 hover:text-blue-800 underline"
                    >
                      Google AI Studio
                    </a>
                    .
                  </p>
                </div>
                
                <Field label="Gemini API Key">
                  <div className="flex gap-2">
                    <Input
                      type={showApiKey ? "text" : "password"}
                      value={apiKey}
                      onChange={(e) => setApiKey(e.target.value)}
                      placeholder="Enter your Gemini API key..."
                      className="flex-1"
                    />
                    <Button
                      icon={showApiKey ? <EyeOff24Regular /> : <Eye24Regular />}
                      appearance="subtle"
                      onClick={() => setShowApiKey(!showApiKey)}
                      title={showApiKey ? "Hide API key" : "Show API key"}
                    />
                  </div>
                </Field>

                {currentApiKey && (
                  <div className="text-xs text-gray-500">
                    Current key: {maskedApiKey}
                  </div>
                )}

                <div className="bg-blue-50 dark:bg-blue-900/20 p-3 rounded-md">
                  <p className="text-xs text-blue-800 dark:text-blue-200">
                    <strong>Security Note:</strong> Your API key is stored securely using Office's built-in storage
                    and is only used to make requests to Google's Gemini API. It's never shared with third parties.
                  </p>
                </div>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={handleClose}>
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={handleSave}
                disabled={isSaving || !apiKey.trim()}
              >
                {isSaving ? <Spinner size="tiny" /> : "Save"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
};

export default ApiKeyManager;
