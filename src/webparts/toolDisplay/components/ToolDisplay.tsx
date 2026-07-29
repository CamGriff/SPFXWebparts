import * as React from "react";
import { useState, useEffect } from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Icon } from "@fluentui/react/lib/Icon";
import { IToolDisplayProps } from "./IToolDisplayProps";
import { ToolsService } from "../../../services/ToolsService";
import { ITool } from "../../../services/models/ITool";

const ToolDisplay: React.FC<IToolDisplayProps> = ({ context, title, columns }) => {

  const [tools, setTools] = useState<ITool[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const service = new ToolsService(context);

  // ─── LOAD ──────────────────────────────────────────────────────

  useEffect(() => {
    const load = async (): Promise<void> => {
      try {
        const [allTools, prefs] = await Promise.all([
          service.getAllTools(),
          service.getUserPreference(),
        ]);

        if (!prefs || prefs.tools.length === 0) {
          // User has no preferences saved yet — show all tools
          setTools(allTools);
        } else {
          // Return tools in the user's saved order
          const ordered = prefs.tools
            .map((id) => allTools.find((t) => t.id === id))
            .filter(Boolean) as ITool[];
          setTools(ordered);
        }
      } catch {
        setError("Error loading your tools. Please try again.");
      } finally {
        setLoading(false);
      }
    };
    void load();
  }, []);

  // ─── RENDER ────────────────────────────────────────────────────

  if (loading) {
    return <Spinner size={SpinnerSize.large} label="Loading your tools..." />;
  }

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  if (tools.length === 0) {
    return (
      <MessageBar messageBarType={MessageBarType.info}>
        No tools found. Visit the My Tools settings page to select your tools.
      </MessageBar>
    );
  }

  const gridStyle: React.CSSProperties = {
    display: "grid",
    gridTemplateColumns: `repeat(${columns}, 1fr)`,
    gap: "16px",
    padding: "16px 0",
  };

  const tileStyle: React.CSSProperties = {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "8px",
    padding: "16px",
    background: "#ffffff",
    borderRadius: "8px",
    boxShadow: "0 1px 4px rgba(0,0,0,0.1)",
    cursor: "pointer",
    textDecoration: "none",
    color: "#323130",
    transition: "box-shadow 0.2s ease",
  };

  return (
    <div style={{ padding: "0 16px" }}>

      {title && (
        <h2 style={{ fontFamily: "Segoe UI", fontWeight: 600, marginBottom: "4px" }}>
          {title}
        </h2>
      )}

      <div style={gridStyle}>
        {tools.map((tool) => (
          <a
            key={tool.id}
            href={tool.url}
            target="_blank"
            rel="noreferrer"
            style={tileStyle}
            onMouseEnter={(e) => {
              (e.currentTarget as HTMLElement).style.boxShadow = "0 4px 12px rgba(0,0,0,0.15)";
            }}
            onMouseLeave={(e) => {
              (e.currentTarget as HTMLElement).style.boxShadow = "0 1px 4px rgba(0,0,0,0.1)";
            }}
          >
            <Icon
              iconName={tool.toolIcon}
              styles={{ root: { fontSize: 28, color: "#0078d4" } }}
            />
            <span style={{ fontSize: 13, fontWeight: 500, textAlign: "center" }}>
              {tool.title}
            </span>
          </a>
        ))}
      </div>

    </div>
  );
};

export default ToolDisplay;