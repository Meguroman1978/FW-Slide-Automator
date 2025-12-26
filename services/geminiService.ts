
import { GoogleGenAI, Type } from "@google/genai";
import { MappingField, Template } from "../types";

export const suggestMappings = async (
  sourceHeaders: string[],
  template: Template
): Promise<MappingField[]> => {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  const prompt = `
    I have a JSON dataset from Sigma with these headers: ${JSON.stringify(sourceHeaders)}.
    The target report template "${template.name}" has these fields: ${JSON.stringify(template.fields)}.
    
    Mapping Strategy:
    - Date: "Event Date Jp", "Started At Jp", "Date"
    - UU (Live): "Unique Viewers (Live)", "ライブ視聴者数", "Unique Viewers Count"
    - UU (Archive): "Unique Viewers (Archive)", "アーカイブ視聴者数", "Unique Viewers Count", "Total Unique Viewers"
    - Peak: "Peak Concurrent Viewers", "最大同時接続数", "Peak Concurrent Viewers Count"
    - Avg Minutes: "Average Watched Minutes", "平均視聴分数"
    - Engagement: "Likes Count", "Reactions", "Chat Count", "Total Chats", "いいね数", "チャット数"
    - Status: "Live Stream Status", "ステータス" (Look for values like 'active' or 'ended')
    - Clicks: "Product Clicks", "Total Product Clicks", "商品クリック数"
    - Time: "経過時間 (分)"

    Please map each target field to the most appropriate source header. 
    Source headers may be in English or Japanese.
    Multiple target fields can map to the same source header if necessary.
    Provide a confidence score (0 to 1).
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              targetKey: { type: Type.STRING },
              sourceKey: { type: Type.STRING },
              confidence: { type: Type.NUMBER }
            },
            required: ["targetKey", "sourceKey", "confidence"]
          }
        }
      }
    });

    const suggestions = JSON.parse(response.text);
    return template.fields.map(tf => {
      const suggestion = suggestions.find((s: any) => s.targetKey === tf.targetKey);
      return {
        targetKey: tf.targetKey,
        label: tf.label,
        sourceKey: suggestion?.sourceKey || "",
        confidence: suggestion?.confidence || 0
      };
    });
  } catch (error) {
    console.error("Gemini mapping suggestion failed:", error);
    return template.fields.map(tf => ({
      targetKey: tf.targetKey,
      label: tf.label,
      sourceKey: "",
      confidence: 0
    }));
  }
};

/**
 * Generates dynamic insight text based on the provided slide data.
 */
export const generateInsight = async (data: any, context: string): Promise<string> => {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  const dataSummary = JSON.stringify(data);

  const prompt = `
    Analyze the following data presented on a business slide and provide a professional, concise "Insight" or "Commentary" in Japanese.
    
    【重要ルール】
    - 提供された「Data」に含まれる数値や項目のみについて言及してください。
    - 時系列データ（経過時間）が含まれる場合は、どの時間帯に何が起きたか具体的に触れてください。
    - 200文字以内で作成してください。
    - 「〜が推測される」「〜の対策が有効」といった、次のアクションにつながる示唆を含めてください。
    - 出力は日本語のテキストのみにしてください。

    文脈: ${context}
    Data: ${dataSummary}
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
    });
    return response.text.trim();
  } catch (error) {
    console.error("Gemini insight generation failed:", error);
    return "データに基づく傾向を分析中です。";
  }
};
