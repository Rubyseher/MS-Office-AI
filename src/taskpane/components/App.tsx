import {
  Button,
  FluentProvider,
  Input,
  Spinner,
  Toast,
  ToastBody,
  Toaster,
  ToastTitle,
  useId,
  useToastController,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import { GoogleGenerativeAI } from "@google/generative-ai";
import * as React from "react";
import { useEffect } from "react";
import {
  insertText,
  registerSelectionChangeHandler,
  removeSelectionChangeHandler,
} from "../taskpane";
import useDarkMode from "./useDarkMode";
import ApiKeyManager from "./ApiKeyManager";
import { getStoredApiKey } from "../utils/apiKeyUtils";

// Initialize with fallback API key, will be updated dynamically
let genAI: GoogleGenerativeAI;
let model: any;

const initializeAI = (apiKey: string) => {
  if (apiKey) {
    genAI = new GoogleGenerativeAI(apiKey);
    model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
  }
};

// Try to initialize with environment API key as fallback
const envApiKey = typeof process !== 'undefined' && process.env ? process.env.REACT_APP_GEMINI_API_KEY : '';
if (envApiKey) {
  initializeAI(envApiKey);
}

const SYSTEM_PROMPT = `You operate within Excel. All responses must be in the following format:
{
  "INSERT_VALUES": a row-major 2D array format,
  "INSERT_ADDRESS": a string format for Excel,
  "RESPONSE_TO_HUMAN": human-readable description of the content
}

Here is the currently selected range: 
`;

const App = () => {
  const [prompt, setPrompt] = React.useState("");
  const [response, setResponse] = React.useState("");
  const [selectedRange, setSelectedRange] = React.useState({
    values: [[]],
    address: "",
  });
  const [loading, setLoading] = React.useState(false);
  const [apiKey, setApiKey] = React.useState("");
  const [isApiKeyLoaded, setIsApiKeyLoaded] = React.useState(false);
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const { isDarkMode } = useDarkMode();

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

  const handlePromptSubmit = async () => {
    if (!apiKey) {
      showToast("Error", "Please set up your API key in settings first", "error");
      return;
    }

    if (!model) {
      showToast("Error", "AI model not initialized. Please check your API key", "error");
      return;
    }

    setLoading(true);
    try {
      const result = await model.generateContent({
        contents: [
          {
            parts: [
              {
                text: SYSTEM_PROMPT + JSON.stringify(selectedRange),
              },
              {
                text: prompt,
              },
            ],
            role: "user",
          },
        ],
      });
      const generatedText = result.response.text() || "No response received.";
      const parsedResponse = getJSONFromResponse(generatedText);
      setResponse(parsedResponse?.RESPONSE_TO_HUMAN || generatedText);

      await insertText(parsedResponse?.INSERT_ADDRESS || "", parsedResponse?.INSERT_VALUES || [[]])
        .then(() => {
          showToast("Merlin's Beard!", "It worked! Now what?", "success");
        })
        .catch((error) => {
          showToast("Ughhhhhhh", error.toString(), "error");
        });
    } catch (error) {
      console.error("Error generating text:", error);
      if (error.toString().includes("API_KEY") || error.toString().includes("401")) {
        showToast("Error", "Invalid API key. Please check your settings.", "error");
      } else {
        showToast("Error", "Failed to generate text", "error");
      }
    } finally {
      setLoading(false);
    }
  };

  const selectedRangeDescription = React.useMemo(() => {
    const match = selectedRange.address.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (match) {
      const [, startCol, startRow, endCol, endRow] = match;
      const cols = endCol.charCodeAt(0) - startCol.charCodeAt(0) + 1;
      const rows = parseInt(endRow) - parseInt(startRow) + 1;
      return `${cols} cols x ${rows} rows selected`;
    }
    return "No range selected";
  }, [selectedRange.address]);

  const handleApiKeyChange = (newApiKey: string) => {
    setApiKey(newApiKey);
    initializeAI(newApiKey);
  };

  // Load stored API key on component mount
  useEffect(() => {
    const loadApiKey = async () => {
      try {
        const storedApiKey = await getStoredApiKey();
        if (storedApiKey) {
          setApiKey(storedApiKey);
          initializeAI(storedApiKey);
        } else {
          // Use environment key as fallback (safely)
          const envKey = typeof process !== 'undefined' && process.env ? process.env.REACT_APP_GEMINI_API_KEY : '';
          if (envKey) {
            setApiKey(envKey);
            initializeAI(envKey);
          }
        }
      } catch (error) {
        console.error("Error loading API key:", error);
      } finally {
        setIsApiKeyLoaded(true);
      }
    };

    loadApiKey();
  }, []);

  useEffect(() => {
    registerSelectionChangeHandler(setSelectedRange);

    return () => {
      removeSelectionChangeHandler();
    };
  }, []);

  return (
    <FluentProvider theme={!isDarkMode ? webLightTheme : webDarkTheme}>
      <div className="flex flex-col h-screen p-4 gap-2">
        {/* Header with settings */}
        <div className="flex justify-between items-center mb-2">
          <h1 className="text-lg font-semibold text-gray-800 dark:text-white">
            Excel AI Assistant
          </h1>
          <ApiKeyManager 
            onApiKeyChange={handleApiKeyChange}
            currentApiKey={apiKey}
          />
        </div>

        {!isApiKeyLoaded ? (
          <div className="flex items-center justify-center h-full">
            <Spinner size="medium" label="Loading..." />
          </div>
        ) : !apiKey ? (
          <div className="flex flex-col items-center justify-center h-auto text-center p-6 my-auto bg-white dark:bg-gray-900 rounded-lg shadow-md">
            <h2 className="text-xl font-bold text-gray-800 dark:text-white mb-4">
               Welcome to [Cheesy AI Generated Name Here]
            </h2>
            <p className="text-gray-600 dark:text-gray-200 mb-6">
              To get started, you need to set up your Google Gemini API key.
            </p>
            <p className="text-sm text-gray-500 dark:text-gray-300 mb-6">
              You can get a free API key from{" "}
              <a
                href="https://aistudio.google.com/app/apikey"
                target="_blank"
                rel="noopener noreferrer"
                className="text-blue-600 hover:text-blue-800 underline"
              >
                Google AI Studio
              </a>
            </p>
            <div className="mb-4">
              <ApiKeyManager 
                onApiKeyChange={handleApiKeyChange}
                currentApiKey={apiKey}
              />
            </div>
          </div>
        ) : !response ? (
          <div className="flex flex-col items-center justify-center h-auto text-center p-6 my-auto bg-white dark:bg-gray-900 rounded-lg shadow-md">
            <h2 className="text-xl font-bold text-gray-800 dark:text-white mb-4">
              Ready to Go! 🚀
            </h2>
            <p className="text-gray-600 dark:text-gray-200 mb-6">Select a range and ask AI to:</p>
            <ul className="text-gray-600 dark:text-gray-200 list-disc list-inside mb-6 text-start">
              <li>Analyze your data</li>
              <li>Provide insights or summaries</li>
              <li>Generate formulas</li>
              <li>Suggest improvements</li>
              <li>[Other AI Generated things AI can do]</li>
            </ul>
            <p className="text-gray-500 dark:text-gray-300 italic">
              Type your question below to get started!
            </p>
          </div>
        ) : (
          <div className="grow">
            <span className="text-xs capitalize text-gray-500 dark:text-gray-300 tracking-wide">
              AI Overlord Says:
            </span>
            <div className="response-container bg-gray-50 dark:bg-gray-800 p-3 rounded-md mt-1">
              <p className="response-text text-sm">{response}</p>
            </div>
          </div>
        )}
        
        {apiKey && (
          <div className="flex gap-2 self-end mt-2">
            <div className="grow">
              <p className="text-xs text-gray-500 dark:text-gray-400 mb-1">
                {selectedRangeDescription}
              </p>
              <Input
                type="text"
                value={prompt}
                onChange={(e) => setPrompt(e.target.value)}
                placeholder="Ask AI about your data..."
                className="w-full"
                onKeyDown={(e) => {
                  if (e.key === "Enter" && !e.shiftKey) {
                    e.preventDefault();
                    handlePromptSubmit();
                  }
                }}
              />
            </div>
            <Button
              onClick={handlePromptSubmit}
              className="self-end shrink-0"
              disabled={loading || !prompt.trim()}
              appearance="primary"
            >
              {loading ? <Spinner size="tiny" /> : "Ask"}
            </Button>
          </div>
        )}
      </div>
      <Toaster
        toasterId={toasterId}
        offset={{
          vertical: 128,
        }}
      />
    </FluentProvider>
  );
};

function getJSONFromResponse(response: string): any {
  try {
    const cleanedResponse = response
      .replace(/^```(?:json)?\s*/i, "") // Handle optional "json" after ```
      .replace(/\s*\n*```$/, "") // Handle closing ```
      .replace(/^[^{]+/, "") // Remove any text before the first `{`
      .replace(/[^}]+$/, "") // Remove any text after the last `}`
      .trim();

    // Validate if the cleaned response starts and ends with curly braces
    if (!cleanedResponse.startsWith("{") || !cleanedResponse.endsWith("}")) {
      throw new Error("Response does not appear to be valid JSON.");
    }
    return JSON.parse(cleanedResponse);
  } catch (error) {
    console.error("Error parsing JSON:", error, "Response:", response);
    return {
      INSERT_VALUES: [[]],
      INSERT_ADDRESS: "",
      RESPONSE_TO_HUMAN: "An error occurred while parsing the AI response.",
    };
  }
}

export default App;
