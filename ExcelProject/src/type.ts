// Tipos para el proyecto ExcelProject
export interface User {
  id?: number; // Opcional, porque en mongoose no está, pero en SQL sí
  username: string;
  email: string;
  password: string;
  created_at?: string; // Opcional, solo en SQL
}

export interface ExcelData {
  id?: number;
  data: any; // JSONB, puede ser cualquier estructura
  created_at?: string;
}

// Respuestas de la API
export interface ApiResponse {
  message: string;
  error?: string;
}

export interface RegisterResponse extends ApiResponse {
  userId?: number;
}

export interface LoginResponse extends ApiResponse {}

export interface SaveExcelResponse extends ApiResponse {}

// ...otros tipos auxiliares si se requieren
