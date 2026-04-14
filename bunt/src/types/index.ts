export interface Student {
  id: string
  name: string
  created_at: string
}

export interface ParentAccount {
  id: string
  student_id: string
  display_name: string
}

export interface Artwork {
  id: string
  student_id: string
  title: string
  image_path: string
  created_at: string
}
