import * as React from "react";
import { useSortable } from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import { Icon } from "@fluentui/react/lib/Icon";
import { ITool } from "../../../services/models/ITool";

export interface ISortableToolCardProps {
  tool: ITool;
  onRemove: (id: string) => void;
}

const SortableToolCard: React.FC<ISortableToolCardProps> = ({ tool, onRemove }) => {

  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: tool.id });

  const style: React.CSSProperties = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.5 : 1,
    display: "flex",
    alignItems: "center",
    gap: "12px",
    padding: "10px 14px",
    marginBottom: "8px",
    background: "#ffffff",
    borderRadius: "6px",
    boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
    cursor: isDragging ? "grabbing" : "grab",
  };

  return (
    <div ref={setNodeRef} style={style} {...attributes} {...listeners}>
      {/* Drag handle icon */}
      <Icon iconName="GripperDotsVertical" styles={{ root: { color: "#999", fontSize: 16 } }} />

      {/* Tool icon */}
      <Icon iconName={tool.toolIcon} styles={{ root: { fontSize: 20, color: "#0078d4" } }} />

      {/* Tool title */}
      <span style={{ flex: 1, fontWeight: 500 }}>{tool.title}</span>

      {/* Remove button */}
      <Icon
        iconName="ChromeClose"
        styles={{ root: { color: "#999", fontSize: 12, cursor: "pointer" } }}
        onClick={(e) => {
          e.stopPropagation();
          onRemove(tool.id);
        }}
      />
    </div>
  );
};

export default SortableToolCard;