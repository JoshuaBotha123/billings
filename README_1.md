# GBS Billings

Internal billings system for Global Business Solutions consultants, finance and leadership.

- **index.html** — sign in, create account, reset password
- **my-billings.html** — consultants capture weekly billings, submit weeks and the month-end final
- **dashboard.html** — personal and company dashboards, targets and projections
- **finance.html** — finance & admin: submissions tracker, invoicing, alerts, people, lists, settings

Data lives in Supabase (Postgres with Row Level Security). The key in `config.js` is the
public "anon" key, which is designed to be public — what each person can see and change is
enforced by the database, not by this website. No passwords or secret keys are stored here.

Internal use only. Access is limited to @globalbusiness.co.za accounts.
