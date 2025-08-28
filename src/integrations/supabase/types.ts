export type Json =
  | string
  | number
  | boolean
  | null
  | { [key: string]: Json | undefined }
  | Json[]

export type Database = {
  // Allows to automatically instantiate createClient with right options
  // instead of createClient<Database, { PostgrestVersion: 'XX' }>(URL, KEY)
  __InternalSupabase: {
    PostgrestVersion: "13.0.4"
  }
  public: {
    Tables: {
      alembic_version: {
        Row: {
          version_num: string
        }
        Insert: {
          version_num: string
        }
        Update: {
          version_num?: string
        }
        Relationships: []
      }
      employee: {
        Row: {
          access_key: string
          created_at: string | null
          default_refinery: string | null
          id: number
          name: string | null
          real_name: string | null
          role: string | null
          updated_at: string | null
          username: string | null
          weekly_goal: number | null
        }
        Insert: {
          access_key: string
          created_at?: string | null
          default_refinery?: string | null
          id?: number
          name?: string | null
          real_name?: string | null
          role?: string | null
          updated_at?: string | null
          username?: string | null
          weekly_goal?: number | null
        }
        Update: {
          access_key?: string
          created_at?: string | null
          default_refinery?: string | null
          id?: number
          name?: string | null
          real_name?: string | null
          role?: string | null
          updated_at?: string | null
          username?: string | null
          weekly_goal?: number | null
        }
        Relationships: []
      }
      entry: {
        Row: {
          created_at: string | null
          date: string
          employee_id: number
          id: number
          observations: string
          points: number
          refinery: string
          updated_at: string | null
        }
        Insert: {
          created_at?: string | null
          date: string
          employee_id: number
          id?: number
          observations: string
          points: number
          refinery: string
          updated_at?: string | null
        }
        Update: {
          created_at?: string | null
          date?: string
          employee_id?: number
          id?: number
          observations?: string
          points?: number
          refinery?: string
          updated_at?: string | null
        }
        Relationships: [
          {
            foreignKeyName: "entry_employee_id_fkey"
            columns: ["employee_id"]
            isOneToOne: false
            referencedRelation: "employee"
            referencedColumns: ["id"]
          },
          {
            foreignKeyName: "fk_entry_employee"
            columns: ["employee_id"]
            isOneToOne: false
            referencedRelation: "employee"
            referencedColumns: ["id"]
          },
        ]
      }
      funcionario: {
        Row: {
          chave_acesso: string
          id: number
          meta_semanal: number | null
          nome: string
        }
        Insert: {
          chave_acesso: string
          id?: number
          meta_semanal?: number | null
          nome: string
        }
        Update: {
          chave_acesso?: string
          id?: number
          meta_semanal?: number | null
          nome?: string
        }
        Relationships: []
      }
      month_reset: {
        Row: {
          id: number
          reset_date: string
        }
        Insert: {
          id?: number
          reset_date: string
        }
        Update: {
          id?: number
          reset_date?: string
        }
        Relationships: []
      }
      notification: {
        Row: {
          created_at: string | null
          data: Json | null
          expires_at: string | null
          id: number
          message: string
          priority: string | null
          read: boolean | null
          type: string
          user_id: number
        }
        Insert: {
          created_at?: string | null
          data?: Json | null
          expires_at?: string | null
          id?: number
          message: string
          priority?: string | null
          read?: boolean | null
          type: string
          user_id: number
        }
        Update: {
          created_at?: string | null
          data?: Json | null
          expires_at?: string | null
          id?: number
          message?: string
          priority?: string | null
          read?: boolean | null
          type?: string
          user_id?: number
        }
        Relationships: [
          {
            foreignKeyName: "notification_user_id_fkey"
            columns: ["user_id"]
            isOneToOne: false
            referencedRelation: "employee"
            referencedColumns: ["id"]
          },
        ]
      }
      notification_preference: {
        Row: {
          desktop_notifications: boolean | null
          email_notifications: boolean | null
          id: number
          notification_types: Json | null
          sound_enabled: boolean | null
          user_id: number
          vibration_enabled: boolean | null
        }
        Insert: {
          desktop_notifications?: boolean | null
          email_notifications?: boolean | null
          id?: number
          notification_types?: Json | null
          sound_enabled?: boolean | null
          user_id: number
          vibration_enabled?: boolean | null
        }
        Update: {
          desktop_notifications?: boolean | null
          email_notifications?: boolean | null
          id?: number
          notification_types?: Json | null
          sound_enabled?: boolean | null
          user_id?: number
          vibration_enabled?: boolean | null
        }
        Relationships: [
          {
            foreignKeyName: "notification_preference_user_id_fkey"
            columns: ["user_id"]
            isOneToOne: false
            referencedRelation: "employee"
            referencedColumns: ["id"]
          },
        ]
      }
      points: {
        Row: {
          created_at: string | null
          description: string | null
          id: number
          refinery_id: number
          updated_at: string | null
          user_id: number
          value: number
        }
        Insert: {
          created_at?: string | null
          description?: string | null
          id?: number
          refinery_id: number
          updated_at?: string | null
          user_id: number
          value: number
        }
        Update: {
          created_at?: string | null
          description?: string | null
          id?: number
          refinery_id?: number
          updated_at?: string | null
          user_id?: number
          value?: number
        }
        Relationships: [
          {
            foreignKeyName: "points_refinery_id_fkey"
            columns: ["refinery_id"]
            isOneToOne: false
            referencedRelation: "refineries"
            referencedColumns: ["id"]
          },
          {
            foreignKeyName: "points_user_id_fkey"
            columns: ["user_id"]
            isOneToOne: false
            referencedRelation: "users"
            referencedColumns: ["id"]
          },
        ]
      }
      refineries: {
        Row: {
          active: boolean | null
          capacity: number | null
          created_at: string | null
          id: number
          location: string | null
          name: string
          updated_at: string | null
        }
        Insert: {
          active?: boolean | null
          capacity?: number | null
          created_at?: string | null
          id?: number
          location?: string | null
          name: string
          updated_at?: string | null
        }
        Update: {
          active?: boolean | null
          capacity?: number | null
          created_at?: string | null
          id?: number
          location?: string | null
          name?: string
          updated_at?: string | null
        }
        Relationships: []
      }
      registro: {
        Row: {
          data: string
          funcionario_id: number
          id: number
          observacoes: string
          pontos: number
          refinaria: string
        }
        Insert: {
          data: string
          funcionario_id: number
          id?: number
          observacoes: string
          pontos: number
          refinaria: string
        }
        Update: {
          data?: string
          funcionario_id?: number
          id?: number
          observacoes?: string
          pontos?: number
          refinaria?: string
        }
        Relationships: [
          {
            foreignKeyName: "registro_funcionario_id_fkey"
            columns: ["funcionario_id"]
            isOneToOne: false
            referencedRelation: "funcionario"
            referencedColumns: ["id"]
          },
        ]
      }
      users: {
        Row: {
          created_at: string | null
          department: string | null
          email: string
          id: number
          is_active: boolean | null
          name: string | null
          password_hash: string | null
          role: string | null
          updated_at: string | null
          username: string
        }
        Insert: {
          created_at?: string | null
          department?: string | null
          email: string
          id?: number
          is_active?: boolean | null
          name?: string | null
          password_hash?: string | null
          role?: string | null
          updated_at?: string | null
          username: string
        }
        Update: {
          created_at?: string | null
          department?: string | null
          email?: string
          id?: number
          is_active?: boolean | null
          name?: string | null
          password_hash?: string | null
          role?: string | null
          updated_at?: string | null
          username?: string
        }
        Relationships: []
      }
    }
    Views: {
      [_ in never]: never
    }
    Functions: {
      [_ in never]: never
    }
    Enums: {
      [_ in never]: never
    }
    CompositeTypes: {
      [_ in never]: never
    }
  }
}

type DatabaseWithoutInternals = Omit<Database, "__InternalSupabase">

type DefaultSchema = DatabaseWithoutInternals[Extract<keyof Database, "public">]

export type Tables<
  DefaultSchemaTableNameOrOptions extends
    | keyof (DefaultSchema["Tables"] & DefaultSchema["Views"])
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
        DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? (DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"] &
      DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Views"])[TableName] extends {
      Row: infer R
    }
    ? R
    : never
  : DefaultSchemaTableNameOrOptions extends keyof (DefaultSchema["Tables"] &
        DefaultSchema["Views"])
    ? (DefaultSchema["Tables"] &
        DefaultSchema["Views"])[DefaultSchemaTableNameOrOptions] extends {
        Row: infer R
      }
      ? R
      : never
    : never

export type TablesInsert<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Insert: infer I
    }
    ? I
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Insert: infer I
      }
      ? I
      : never
    : never

export type TablesUpdate<
  DefaultSchemaTableNameOrOptions extends
    | keyof DefaultSchema["Tables"]
    | { schema: keyof DatabaseWithoutInternals },
  TableName extends DefaultSchemaTableNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"]
    : never = never,
> = DefaultSchemaTableNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaTableNameOrOptions["schema"]]["Tables"][TableName] extends {
      Update: infer U
    }
    ? U
    : never
  : DefaultSchemaTableNameOrOptions extends keyof DefaultSchema["Tables"]
    ? DefaultSchema["Tables"][DefaultSchemaTableNameOrOptions] extends {
        Update: infer U
      }
      ? U
      : never
    : never

export type Enums<
  DefaultSchemaEnumNameOrOptions extends
    | keyof DefaultSchema["Enums"]
    | { schema: keyof DatabaseWithoutInternals },
  EnumName extends DefaultSchemaEnumNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"]
    : never = never,
> = DefaultSchemaEnumNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[DefaultSchemaEnumNameOrOptions["schema"]]["Enums"][EnumName]
  : DefaultSchemaEnumNameOrOptions extends keyof DefaultSchema["Enums"]
    ? DefaultSchema["Enums"][DefaultSchemaEnumNameOrOptions]
    : never

export type CompositeTypes<
  PublicCompositeTypeNameOrOptions extends
    | keyof DefaultSchema["CompositeTypes"]
    | { schema: keyof DatabaseWithoutInternals },
  CompositeTypeName extends PublicCompositeTypeNameOrOptions extends {
    schema: keyof DatabaseWithoutInternals
  }
    ? keyof DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"]
    : never = never,
> = PublicCompositeTypeNameOrOptions extends {
  schema: keyof DatabaseWithoutInternals
}
  ? DatabaseWithoutInternals[PublicCompositeTypeNameOrOptions["schema"]]["CompositeTypes"][CompositeTypeName]
  : PublicCompositeTypeNameOrOptions extends keyof DefaultSchema["CompositeTypes"]
    ? DefaultSchema["CompositeTypes"][PublicCompositeTypeNameOrOptions]
    : never

export const Constants = {
  public: {
    Enums: {},
  },
} as const
