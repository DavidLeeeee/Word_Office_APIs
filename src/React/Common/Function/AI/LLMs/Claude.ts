export enum Claude_Model {
  Claude_3_5_Sonnet = "claude-3-5-sonnet-20241022",
  Claude_3_Opus = "claude-3-opus-20240229",
  Claude_3_Sonnet = "claude-3-sonnet-20240229",
}

class Claude {
  async generateText(prompt: string, model: Claude_Model = Claude_Model.Claude_3_5_Sonnet): Promise<string> {
    // Claude API 호출 구현
    // TODO: 실제 API 구현
    return `Claude ${model} response for: ${prompt}`;
  }

  async generateStream(prompt: string, model: Claude_Model = Claude_Model.Claude_3_5_Sonnet): Promise<ReadableStream> {
    // Claude Stream API 호출 구현
    // TODO: 실제 API 구현
    throw new Error("Not implemented");
  }
}

export default Claude;
