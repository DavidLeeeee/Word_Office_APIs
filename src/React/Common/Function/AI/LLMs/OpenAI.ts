export enum OpenAI_Model {
  GPT_4o = "gpt-4o",
  GPT_4_Turbo = "gpt-4-turbo",
  GPT_4 = "gpt-4",
  GPT_3_5_Turbo = "gpt-3.5-turbo",
}

interface OpenAIResponse {
  choices: Array<{
    message: {
      content: string;
    };
  }>;
}

class OpenAI {
  private apiKey: string;

  constructor() {
    const apiKey = process.env.OPENAI_KEY;
    if (!apiKey) {
      throw new Error("OPENAI_KEY 환경 변수가 설정되지 않았습니다. .env 파일에 OPENAI_KEY를 추가해주세요.");
    }
    this.apiKey = apiKey;
  }

  async generateText(prompt: string, model: OpenAI_Model = OpenAI_Model.GPT_4o): Promise<string> {
    try {
      const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${this.apiKey}`,
        },
        body: JSON.stringify({
          model: model,
          messages: [
            {
              role: "user",
              content: prompt,
            },
          ],
        }),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(`OpenAI API 오류: ${response.status} ${response.statusText} - ${JSON.stringify(errorData)}`);
      }

      const data: OpenAIResponse = await response.json();
      
      if (!data.choices || data.choices.length === 0) {
        throw new Error("OpenAI API 응답에 선택지가 없습니다.");
      }

      return data.choices[0].message.content;
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`OpenAI API 호출 실패: ${error.message}`);
      }
      throw new Error("OpenAI API 호출 중 알 수 없는 오류가 발생했습니다.");
    }
  }

  async generateStream(prompt: string, model: OpenAI_Model = OpenAI_Model.GPT_4o): Promise<ReadableStream> {
    // Stream은 당장 필요 없으므로 구현하지 않음
    throw new Error("Stream 기능은 아직 구현되지 않았습니다.");
  }
}

export default OpenAI;
