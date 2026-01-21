import OpenAI, { OpenAI_Model } from "./LLMs/OpenAI";
import Gemini, { Gemini_Model } from "./LLMs/Gemini";
import Claude, { Claude_Model } from "./LLMs/Claude";
import Gemma, { Gemma_Model } from "./LLMs/Gemma";
import Qwen, { Qwen_Model } from "./LLMs/Qwen";

// 모델 인스턴스 생성
const openai = new OpenAI();
const gemini = new Gemini();
const claude = new Claude();
const gemma = new Gemma();
const qwen = new Qwen();

// Factory 객체 - AI_Factory.xxx.xxx 형태로 호출 가능
const AI_Factory = {
  openai: {
    GPT_4o: {
      generateText: (prompt: string) => openai.generateText(prompt, OpenAI_Model.GPT_4o),
      generateStream: (prompt: string) => openai.generateStream(prompt, OpenAI_Model.GPT_4o),
    },
    GPT_4_Turbo: {
      generateText: (prompt: string) => openai.generateText(prompt, OpenAI_Model.GPT_4_Turbo),
      generateStream: (prompt: string) => openai.generateStream(prompt, OpenAI_Model.GPT_4_Turbo),
    },
    GPT_4: {
      generateText: (prompt: string) => openai.generateText(prompt, OpenAI_Model.GPT_4),
      generateStream: (prompt: string) => openai.generateStream(prompt, OpenAI_Model.GPT_4),
    },
    GPT_3_5_Turbo: {
      generateText: (prompt: string) => openai.generateText(prompt, OpenAI_Model.GPT_3_5_Turbo),
      generateStream: (prompt: string) => openai.generateStream(prompt, OpenAI_Model.GPT_3_5_Turbo),
    },
  },
  gemini: {
    GeminiPro: {
      generateText: (prompt: string) => gemini.generateText(prompt, Gemini_Model.GeminiPro),
      generateStream: (prompt: string) => gemini.generateStream(prompt, Gemini_Model.GeminiPro),
    },
    GeminiUltra: {
      generateText: (prompt: string) => gemini.generateText(prompt, Gemini_Model.GeminiUltra),
      generateStream: (prompt: string) => gemini.generateStream(prompt, Gemini_Model.GeminiUltra),
    },
  },
  claude: {
    Claude_3_5_Sonnet: {
      generateText: (prompt: string) => claude.generateText(prompt, Claude_Model.Claude_3_5_Sonnet),
      generateStream: (prompt: string) => claude.generateStream(prompt, Claude_Model.Claude_3_5_Sonnet),
    },
    Claude_3_Opus: {
      generateText: (prompt: string) => claude.generateText(prompt, Claude_Model.Claude_3_Opus),
      generateStream: (prompt: string) => claude.generateStream(prompt, Claude_Model.Claude_3_Opus),
    },
    Claude_3_Sonnet: {
      generateText: (prompt: string) => claude.generateText(prompt, Claude_Model.Claude_3_Sonnet),
      generateStream: (prompt: string) => claude.generateStream(prompt, Claude_Model.Claude_3_Sonnet),
    },
  },
  gemma: {
    Gemma_2B: {
      generateText: (prompt: string) => gemma.generateText(prompt, Gemma_Model.Gemma_2B),
      generateStream: (prompt: string) => gemma.generateStream(prompt, Gemma_Model.Gemma_2B),
    },
    Gemma_7B: {
      generateText: (prompt: string) => gemma.generateText(prompt, Gemma_Model.Gemma_7B),
      generateStream: (prompt: string) => gemma.generateStream(prompt, Gemma_Model.Gemma_7B),
    },
  },
  qwen: {
    Qwen_Turbo: {
      generateText: (prompt: string) => qwen.generateText(prompt, Qwen_Model.Qwen_Turbo),
      generateStream: (prompt: string) => qwen.generateStream(prompt, Qwen_Model.Qwen_Turbo),
    },
    Qwen_Plus: {
      generateText: (prompt: string) => qwen.generateText(prompt, Qwen_Model.Qwen_Plus),
      generateStream: (prompt: string) => qwen.generateStream(prompt, Qwen_Model.Qwen_Plus),
    },
    Qwen_Max: {
      generateText: (prompt: string) => qwen.generateText(prompt, Qwen_Model.Qwen_Max),
      generateStream: (prompt: string) => qwen.generateStream(prompt, Qwen_Model.Qwen_Max),
    },
  },
};

export default AI_Factory;

// 사용 예시:
// AI_Factory.openai.GPT_4o.generateText("Hello, world!")
// AI_Factory.gemini.GeminiPro.generateText("Hello, world!")
// AI_Factory.claude.Claude_3_5_Sonnet.generateText("Hello, world!")
