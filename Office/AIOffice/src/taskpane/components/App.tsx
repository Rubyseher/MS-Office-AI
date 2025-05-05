import { Button, Input, makeStyles } from "@fluentui/react-components";
import { GoogleGenerativeAI } from "@google/generative-ai";
import * as React from "react";
import { insertText } from "../taskpane";

const genAI = new GoogleGenerativeAI(process.env.REACT_APP_GEMINI_API_KEY || "");
const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();
  const [prompt, setPrompt] = React.useState("");
  const [response, setResponse] = React.useState("");

  const handlePromptSubmit = async () => {
    try {
      const result = await model.generateContent({
        contents: [
          {
            parts: [
              {
                text: prompt,
              },
            ],
            role: "user",
          },
        ],
      });
      const generatedText = result.response.text() || "No response received.";
      setResponse(generatedText);
      insertText(generatedText); // Invoke TextInsertion with the response
    } catch (error) {
      console.error("Error generating text:", error);
    }
  };

  return (
    <div className={styles.root}>
      <div style={{ margin: "20px 0" }}>
        <Input
          type="text"
          value={prompt}
          onChange={(e) => setPrompt(e.target.value)}
          placeholder="Enter your prompt here"
          style={{ width: "80%", padding: "10px", fontSize: "16px" }}
        />
        <Button onClick={handlePromptSubmit} style={{ marginLeft: "10px", padding: "10px 20px" }}>
          Submit
        </Button>
      </div>
    </div>
  );
};

export default App;
