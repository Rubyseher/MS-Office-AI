import { makeStyles } from "@fluentui/react-components";
import { DesignIdeas24Regular, LockOpen24Regular, Ribbon24Regular } from "@fluentui/react-icons";
import { GoogleGenerativeAI } from "@google/generative-ai";
import * as React from "react";
import { insertText } from "../taskpane";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";

const genAI = new GoogleGenerativeAI(process.env.API_KEY);
const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
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

  const listItems: HeroListItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      <div style={{ margin: "20px 0" }}>
        <input
          type="text"
          value={prompt}
          onChange={(e) => setPrompt(e.target.value)}
          placeholder="Enter your prompt here"
          style={{ width: "80%", padding: "10px", fontSize: "16px" }}
        />
        <button onClick={handlePromptSubmit} style={{ marginLeft: "10px", padding: "10px 20px" }}>
          Submit
        </button>
      </div>
      <div>
        <strong>Response:</strong>
        <p>{response}</p>
      </div>
      <TextInsertion insertText={insertText} />
    </div>
  );
};

export default App;