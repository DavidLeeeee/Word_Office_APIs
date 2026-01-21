export enum Gemma_Model {
  Gemma_2B = "gemma-2b",
  Gemma_7B = "gemma-7b",
}

class Gemma {
  async generateText(prompt: string, model: Gemma_Model = Gemma_Model.Gemma_7B): Promise<string> {
    // Gemma API 호출 구현
    // TODO: 실제 API 구현
    return `Gemma ${model} response for: ${prompt}`;
  }

  async generateStream(prompt: string, model: Gemma_Model = Gemma_Model.Gemma_7B): Promise<ReadableStream> {
    // Gemma Stream API 호출 구현
    // TODO: 실제 API 구현
    throw new Error("Not implemented");
  }
}

export default Gemma;
