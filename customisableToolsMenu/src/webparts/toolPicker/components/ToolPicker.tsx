import * as React from "react";
import { useState, useEffect } from "react";
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragEndEvent,
} from "@dnd-kit/core";
import {
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  arrayMove,
} from "@dnd-kit/sortable";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { IToolPickerProps } from "./IToolPickerProps";
import { ToolsService } from "../../../services/ToolsService";
import { ITool } from "../../../services/models/ITool";
import SortableToolCard from "./SortableToolCard";
import AvailableToolCard from "./AvailableToolCard";

const ToolPicker: React.FC<IToolPickerProps> = ({ context }) => {

  const [allTools, setAllTools] = useState<ITool[]>([]);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [saving, setSaving] = useState<boolean>(false);
  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);

  const service = new ToolsService(context);

  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  // ─── LOAD ──────────────────────────────────────────────────────

  useEffect(() => {
    const load = async (): Promise<void> => {
      try {
        const [tools, prefs] = await Promise.all([
          service.getAllTools(),
          service.getUserPreference(),
        ]);
        setAllTools(tools);
        setSelectedIds(prefs ? prefs.tools : []);
      } catch (err) {
        setMessage({ text: "Error loading tools. Please try again.", type: MessageBarType.error });
      } finally {
        setLoading(false);
      }
    };
    void load();
  }, []);

  // ─── DRAG END ──────────────────────────────────────────────────

  const handleDragEnd = (event: DragEndEvent): void => {
    const { active, over } = event;
    if (over && active.id !== over.id) {
      setSelectedIds((ids) => {
        const oldIndex = ids.indexOf(active.id as string);
        const newIndex = ids.indexOf(over.id as string);
        return arrayMove(ids, oldIndex, newIndex);
      });
    }
  };

  // ─── ADD / REMOVE ──────────────────────────────────────────────

  const handleAdd = (id: string): void => {
    setSelectedIds((prev) => [...prev, id]);
  };

  const handleRemove = (id: string): void => {
    setSelectedIds((prev) => prev.filter((t) => t !== id));
  };

  // ─── SAVE ──────────────────────────────────────────────────────

  const handleSave = async (): Promise<void> => {
    setSaving(true);
    setMessage(null);
    try {
      await service.saveUserPreference(selectedIds);
      setMessage({ text: "Your tools have been saved!", type: MessageBarType.success });
    } catch (err) {
      setMessage({ text: "Error saving tools. Please try again.", type: MessageBarType.error });
    } finally {
      setSaving(false);
    }
  };

  // ─── DERIVED ───────────────────────────────────────────────────

  const selectedTools = selectedIds
    .map((id) => allTools.find((t) => t.id === id))
    .filter(Boolean) as ITool[];

  const availableTools = allTools.filter((t) => !selectedIds.includes(t.id.toString()));

  // ─── RENDER ────────────────────────────────────────────────────

  if (loading) {
    return <Spinner size={SpinnerSize.large} label="Loading tools..." />;
  }

  return (
    <div style={{ maxWidth: 480, padding: "20px" }}>

      <h2 style={{ marginBottom: "4px" }}>My Tools</h2>
      <p style={{ color: "#605e5c", marginBottom: "20px" }}>
        Select and drag to order your tools.
      </p>

      {message && (
        <MessageBar
          messageBarType={message.type}
          onDismiss={() => setMessage(null)}
          style={{ marginBottom: "16px" }}
        >
          {message.text}
        </MessageBar>
      )}

      {/* ── Selected tools (draggable) ── */}
      {selectedTools.length > 0 && (
        <>
          <p style={{ fontWeight: 600, marginBottom: "8px" }}>
            Selected ({selectedTools.length})
          </p>
          <DndContext
            sensors={sensors}
            collisionDetection={closestCenter}
            onDragEnd={handleDragEnd}
          >
            <SortableContext
              items={selectedIds}
              strategy={verticalListSortingStrategy}
            >
              {selectedTools.map((tool) => (
                <SortableToolCard
                  key={tool.id}
                  tool={tool}
                  onRemove={handleRemove}
                />
              ))}
            </SortableContext>
          </DndContext>
        </>
      )}

      {/* ── Available tools ── */}
      {availableTools.length > 0 && (
        <>
          <p style={{ fontWeight: 600, margin: "16px 0 8px" }}>
            Available ({availableTools.length})
          </p>
          {availableTools.map((tool) => (
            <AvailableToolCard
              key={tool.id}
              tool={tool}
              onAdd={handleAdd}
            />
          ))}
        </>
      )}

      {/* ── Actions ── */}
      <div style={{ display: "flex", gap: "10px", marginTop: "24px" }}>
        <PrimaryButton
          text={saving ? "Saving..." : "Save My Tools"}
          onClick={handleSave}
          disabled={saving}
        />
        <DefaultButton
          text="Reset"
          onClick={() => setSelectedIds([])}
          disabled={saving}
        />
      </div>

    </div>
  );
};

export default ToolPicker;