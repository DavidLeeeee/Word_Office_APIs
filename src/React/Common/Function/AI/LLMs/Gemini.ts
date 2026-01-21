export enum Gemini_Model {
  GeminiPro = "gemini-pro",
  GeminiUltra = "gemini-ultra",
}

class Gemini {
  async generateText(prompt: string, model: Gemini_Model = Gemini_Model.GeminiPro): Promise<string> {
    // Gemini API 호출 구현
    // TODO: 실제 API 구현
    return `Gemini ${model} response for: ${prompt}`;
  }

  async generateStream(prompt: string, model: Gemini_Model = Gemini_Model.GeminiPro): Promise<ReadableStream> {
    // Gemini Stream API 호출 구현
    // TODO: 실제 API 구현
    throw new Error("Not implemented");
  }
}

export default Gemini;
