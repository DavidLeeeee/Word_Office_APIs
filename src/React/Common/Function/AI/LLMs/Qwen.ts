export enum Qwen_Model {
  Qwen_Turbo = "qwen-turbo",
  Qwen_Plus = "qwen-plus",
  Qwen_Max = "qwen-max",
}

class Qwen {
  async generateText(prompt: string, model: Qwen_Model = Qwen_Model.Qwen_Turbo): Promise<string> {
    // Qwen API 호출 구현
    // TODO: 실제 API 구현
    return `Qwen ${model} response for: ${prompt}`;
  }

  async generateStream(prompt: string, model: Qwen_Model = Qwen_Model.Qwen_Turbo): Promise<ReadableStream> {
    // Qwen Stream API 호출 구현
    // TODO: 실제 API 구현
    throw new Error("Not implemented");
  }
}

export default Qwen;
