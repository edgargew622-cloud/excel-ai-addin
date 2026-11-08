export const schemaV2 = `
CREATE TABLE commercial.tasks (
 org_id uuid NOT NULL, id uuid NOT NULL, user_id uuid NOT NULL,
 key_mode text NOT NULL CHECK(key_mode IN ('hosted','byok')),
 state text NOT NULL DEFAULT 'active' CHECK(state IN ('active','completed','cancelled','stopped')),
 max_steps integer NOT NULL CHECK(max_steps BETWEEN 1 AND 50), steps integer NOT NULL DEFAULT 0 CHECK(steps>=0 AND steps<=max_steps),
 cap bigint NOT NULL CHECK(cap>0), spent bigint NOT NULL DEFAULT 0 CHECK(spent>=0), reserved bigint NOT NULL DEFAULT 0 CHECK(reserved>=0),
 created_at timestamptz NOT NULL DEFAULT now(), deadline timestamptz NOT NULL, CHECK(spent+reserved<=cap),
 PRIMARY KEY(org_id,id), FOREIGN KEY(org_id,user_id) REFERENCES commercial.memberships(org_id,user_id)
);
ALTER TABLE commercial.usage_requests ADD COLUMN task_id uuid;
ALTER TABLE commercial.usage_requests ADD FOREIGN KEY(org_id,task_id) REFERENCES commercial.tasks(org_id,id);
CREATE INDEX usage_task ON commercial.usage_requests(org_id,task_id);
CREATE TABLE commercial.auth_flows (
 state_hash text PRIMARY KEY, nonce text NOT NULL, verifier jsonb NOT NULL, parent_challenge text NOT NULL,
 expires_at timestamptz NOT NULL, claimed boolean NOT NULL DEFAULT false
);
CREATE TABLE commercial.auth_redemptions (
 ticket_hash text PRIMARY KEY, user_id uuid NOT NULL REFERENCES commercial.users(id), parent_challenge text NOT NULL, expires_at timestamptz NOT NULL
);
CREATE TABLE commercial.auth_sessions (
 token_hash text PRIMARY KEY, user_id uuid NOT NULL REFERENCES commercial.users(id), expires_at timestamptz NOT NULL, revoked boolean NOT NULL DEFAULT false
);
CREATE INDEX auth_sessions_user ON commercial.auth_sessions(user_id);
`;
