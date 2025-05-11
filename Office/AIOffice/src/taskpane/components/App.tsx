import { Button, Input, makeStyles } from "@fluentui/react-components";
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

interface AppProps {
  title: string;
}

const SYSTEM_PROMPT = `You operate within Excel. All responses must be in the following format:
{
  "INSERT_VALUES": a row-major 2D array format,
  "INSERT_ADDRESS": a string format for Excel,
  "RESPONSE_TO_HUMAN": human-readable description of the content
}

Here is the currently selected range: 
`;

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const [prompt, setPrompt] = React.useState("");
  const [response, setResponse] = React.useState("");
  const [selectedRange, setSelectedRange] = React.useState({
    values: [[]],
    address: "",
  });

  const handlePromptSubmit = async () => {
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
      insertText(parsedResponse?.INSERT_ADDRESS || "", parsedResponse?.INSERT_VALUES || [[]]);
    } catch (error) {
      console.error("Error generating text:", error);
    }
  };

  useEffect(() => {
    registerSelectionChangeHandler(setSelectedRange);

    return () => {
      removeSelectionChangeHandler();
    };
  }, []);

  return (
    <div className={styles.root}>
      <div style={{ margin: "20px 0" }}>
        <p>{response}</p>
        <p className="bg-blue-200 p-2">{JSON.stringify(selectedRange)}</p>
        <Input
          type="text"
          value={prompt}
          onChange={(e) => setPrompt(e.target.value)}
          placeholder="Ask AI"
          className="ml-8"
        />
        <Button onClick={handlePromptSubmit}>Submit</Button>
      </div>
    </div>
  );
};

function getJSONFromResponse(response: string): any {
  try {
    const cleanedResponse = response
      .replace(/^```json\s*/, "")
      .replace(/\s*\n*```$/, "")
      .trim();
    return JSON.parse(cleanedResponse);
  } catch (error) {
    console.error("Error parsing JSON:", error);
    return null;
  }
}

export default App;
