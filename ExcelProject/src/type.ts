// Tipos de API alineados con el backend

// Utilidades JSON seguras (para evitar any)
export type JSONPrimitive = string | number | boolean | null;
export type JSONValue = JSONPrimitive | JSONObject | JSONArray;
export interface JSONObject { [key: string]: JSONValue }
export interface JSONArray extends Array<JSONValue> {}

// Entidades básicas (si se necesitan en el cliente)
export interface User {
  id?: number;
  username: string;
  email: string;
  created_at?: string; // ISO string si viene desde SQL
}

// Forma estándar de error que usa el backend
export interface ErrorResponse {
  error: string;
}

// =====================
// Auth: register/login
// =====================
export interface RegisterRequest {
  username: string;
  email: string;
  password: string;
}

export interface RegisterSuccess {
  message: string; // "Usuario registrado correctamente"
  userId: number;
}

export type RegisterResponse = RegisterSuccess | ErrorResponse;

export interface LoginRequest {
  username: string;
  password: string;
}

export interface LoginSuccess {
  message: string; // "Login exitoso"
}

export type LoginResponse = LoginSuccess | ErrorResponse;

// =====================
// Excel: guardar datos
// =====================
export interface SaveExcelRequest {
  // El backend espera req.body.data
  data: JSONValue;
}

export interface SaveExcelSuccess {
  message: string; // "Datos de Excel guardados correctamente en la base de datos"
}

export type SaveExcelResponse = SaveExcelSuccess | ErrorResponse;

