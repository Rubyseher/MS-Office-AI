import {
  Button,
  Input,
  Spinner,
  Toast,
  ToastBody,
  Toaster,
  ToastTitle,
  useId,
  useToastController,
} from "@fluentui/react-components";
import { GoogleGenerativeAI } from "@google/generative-ai";
import * as React from "react";
import { useEffect } from "react";
import {
  insertText,
  registerSelectionChangeHandler,
  removeSelectionChangeHandler,
} from "../taskpane";

const genAI = new GoogleGenerativeAI(process.env.REACT_APP_GEMINI_API_KEY || "");
const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

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
  const toasterId = useId("toaster");
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

  const handlePromptSubmit = async () => {
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
      showToast("Error", "Failed to generate text", "error");
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

  useEffect(() => {
    registerSelectionChangeHandler(setSelectedRange);

    return () => {
      removeSelectionChangeHandler();
    };
  }, []);

  return (
    <>
      <div className="flex flex-col h-screen bg-gray-50 p-4 gap-2">
        {!response ? (
          <div className="flex flex-col items-center justify-center h-auto text-center p-4 my-auto bg-white rounded-md shadow-md">
            <h1 className="text-2xl font-bold text-gray-800 mb-4">
              Welcome to [Cheesy AI Generated Name Here]
            </h1>
            <p className="text-gray-600 mb-6">Select a range and ask AI to:</p>
            <ul className="text-gray-600 list-disc list-inside mb-6 text-start">
              <li>Analyze your data</li>
              <li>Provide insights or summaries</li>
              <li>Generate formulas</li>
              <li>Suggest improvements</li>
              <li>[Other AI Generated things AI can do]</li>
            </ul>
            <p className="text-gray-500 italic">Type your question below to get started!</p>
          </div>
        ) : (
          <div className="h-full">
            <div className="bg-blue-100 p-4 rounded-lg shadow-md w-full max-w-lg mb-4">
              <p className="text-blue-800 text-sm font-medium">{response}</p>
            </div>
          </div>
        )}
        <div className="flex gap-2 self-end">
          <div className="grow-1">
            <p className="bg-blue-100 p-2 w-full rounded-t-md text-xs text-blue-800 font-semibold capitalize tracking-wide">
              {selectedRangeDescription}
            </p>
            <Input
              type="text"
              value={prompt}
              onChange={(e) => setPrompt(e.target.value)}
              placeholder="Ask AI"
              className="w-full"
            />
          </div>
          <Button
            onClick={handlePromptSubmit}
            className="self-end shrink-0"
            disabled={loading}
            appearance={loading ? "primary" : "secondary"}
          >
            {loading ? <Spinner size="tiny" /> : "Ask"}
          </Button>
        </div>
      </div>
      <Toaster
        toasterId={toasterId}
        offset={{
          vertical: 128,
        }}
      />
    </>
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
