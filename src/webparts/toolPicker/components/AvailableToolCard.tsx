import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";
import { ITool } from "../../../services/models/ITool";

export interface IAvailableToolCardProps {
  tool: ITool;
  onAdd: (id: string) => void;
}

const AvailableToolCard: React.FC<IAvailableToolCardProps> = ({ tool, onAdd }) => {

  const style: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: "12px",
    padding: "10px 14px",
    marginBottom: "8px",
    background: "#f3f2f1",
    borderRadius: "6px",
    cursor: "pointer",
    opacity: 0.8,
  };

  return (
    <div style={style} onClick={() => onAdd(tool.id)}>
      {/* Tool icon */}
      <Icon iconName={tool.toolIcon} styles={{ root: { fontSize: 20, color: "#605e5c" } }} />

      {/* Tool title */}
      <span style={{ flex: 1 }}>{tool.title}</span>

      {/* Add indicator */}
      <Icon
        iconName="Add"
        styles={{ root: { color: "#0078d4", fontSize: 14 } }}
      />
    </div>
  );
};

export default AvailableToolCard;