-- Create transport_plan_dhaka table
CREATE TABLE IF NOT EXISTS transport_plan_dhaka (
    id BIGSERIAL PRIMARY KEY,
    sl_number TEXT,
    checklist_category TEXT,
    oversight_manager TEXT,
    action_item TEXT,
    responsible_person_name TEXT,
    responsible_person_designation TEXT,
    responsible_person_email TEXT,
    reminder_cc_email TEXT,
    due_date DATE,
    reminder_days TEXT,
    reminder_sent TEXT,
    reminder_count INTEGER DEFAULT 0,
    status TEXT,
    comments TEXT,
    sheet_source TEXT,
    last_synced TIMESTAMPTZ DEFAULT NOW(),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Enable Row Level Security but create a policy that allows all operations
ALTER TABLE transport_plan_dhaka ENABLE ROW LEVEL SECURITY;

-- Create policy to allow all operations (you can restrict this later)
CREATE POLICY "Allow all operations" ON transport_plan_dhaka
FOR ALL USING (true);

-- Create index for better performance
CREATE INDEX IF NOT EXISTS idx_transport_plan_dhaka_status ON transport_plan_dhaka(status);
CREATE INDEX IF NOT EXISTS idx_transport_plan_dhaka_due_date ON transport_plan_dhaka(due_date);