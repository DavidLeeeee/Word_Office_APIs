export interface PIIEntity {
  type: string;
  original: string;
  masked: string;
  start: number;
  end: number;
  detection_source: string;
  confidence: number;
  applied_policy: string;
  address_components: any;
}

export interface PIICombinationMatch {
  rule_id: number;
  rule_name: string;
  matched_entities: string[];
  action_type: string;
  risk_level: string;
}

export interface PIIProcessingTime {
  preprocess_ms: number;
  presidio_ms: number;
  kobert_ms: number;
  gemma_ms: number;
  instruction_ms: number;
  classification_ms: number;
  position_estimator_ms: number;
  cascade_ms: number;
  cascade_chunks: number;
  cascade_coverage: number;
  total_ms: number;
}

export interface PIIData {
  session_id: string;
  status: string;
  entities: PIIEntity[];
  masked_text: string;
  original_text: string;
  processing_time: PIIProcessingTime;
  detection_mode: string;
  combination_matches: PIICombinationMatch[];
  metadata: any;
}

export interface PIIResponse {
  result: number;
  message: string;
  data: PIIData;
}

export type DetectionStrategy = "pattern_instruction" | "pattern" | "instruction";

interface PIIFilterRequest {
  text: string;
  user_id: string;
  detection_strategy: DetectionStrategy;
}

class PII_Filter {
  private readonly apiUrl = process.env.NODE_ENV === "development" 
    ? "/api/audit/check" 
    : "https://chat.k-armor.ai:8082/audit/check";

  async check(
    text: string,
    userId: string = "1",
    detectionStrategy: DetectionStrategy = "pattern_instruction"
  ): Promise<PIIResponse> {
    try {
      const requestBody: PIIFilterRequest = {
        text: text,
        user_id: userId,
        detection_strategy: detectionStrategy,
      };

      const response = await fetch(this.apiUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        const errorText = await response.text().catch(() => "");
        throw new Error(`PII Filter API 오류: ${response.status} ${response.statusText} - ${errorText}`);
      }

      const data: PIIResponse = await response.json();

      if (data.result !== 0) {
        throw new Error(`PII Filter API 오류: ${data.message || "알 수 없는 오류"}`);
      }

      return data;
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`PII Filter API 호출 실패: ${error.message}`);
      }
      throw new Error("PII Filter API 호출 중 알 수 없는 오류가 발생했습니다.");
    }
  }

  async getMaskedText(
    text: string,
    userId: string = "1",
    detectionStrategy: DetectionStrategy = "pattern_instruction"
  ): Promise<string> {
    const result = await this.check(text, userId, detectionStrategy);
    const maskedText = result.data?.masked_text;
    const originalText = result.data?.original_text;
    
    // masked_text가 null이면 original_text 사용
    if (!maskedText || typeof maskedText !== "string") {
      if (originalText && typeof originalText === "string") {
        return originalText;
      }
      throw new Error("PII 필터링 결과에서 텍스트를 찾을 수 없습니다.");
    }
    return maskedText;
  }

  async getEntities(
    text: string,
    userId: string = "1",
    detectionStrategy: DetectionStrategy = "pattern_instruction"
  ): Promise<PIIEntity[]> {
    const result = await this.check(text, userId, detectionStrategy);
    return result.data.entities;
  }
}

const piiFilter = new PII_Filter();

export default piiFilter;

// 사용 예시:
// const result = await PII_Filter.check("제 주민번호는 901215-1234567입니다.");
// const maskedText = await PII_Filter.getMaskedText("제 주민번호는 901215-1234567입니다.");
// const entities = await PII_Filter.getEntities("제 주민번호는 901215-1234567입니다.");