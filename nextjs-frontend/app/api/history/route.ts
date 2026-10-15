import { NextResponse } from 'next/server';
import { CosmosClient } from '@azure/cosmos';

// Initialize Cosmos DB Client
const endpoint = process.env.COSMOS_DB_URI || "";
const key = process.env.COSMOS_DB_KEY || "";
const databaseId = process.env.COSMOS_DB_DATABASE || "design-doc-generator";
const containerId = process.env.COSMOS_DB_TASKS_CONTAINER || "tasks";

let client: CosmosClient | null = null;
if (endpoint && key) {
    client = new CosmosClient({ endpoint, key });
}

export const dynamic = 'force-dynamic';

export async function GET(request: Request) {
    if (!client) {
        return NextResponse.json({ error: "Cosmos DB is not configured." }, { status: 500 });
    }

    try {
        const { searchParams } = new URL(request.url);
        const activeOnly = searchParams.get('active_only') === 'true';

        const database = client.database(databaseId);
        const container = database.container(containerId);

        let query = "SELECT c.id, c.taskId, c.filename, c.generated_filename, c.documentType, c.createdAt, c.status, c.cost_metrics, c.cost_summary, c.blobPath, c.storage, c.isDeleted, c.deletedAt FROM c";
        if (activeOnly) {
            query += " WHERE (NOT IS_DEFINED(c.isDeleted) OR c.isDeleted = false)";
        }
        query += " ORDER BY c.createdAt DESC";

        const { resources: items } = await container.items.query({ query }).fetchAll();

        return NextResponse.json({ data: items });
    } catch (error: any) {
        console.error("Cosmos DB Fetch Error (History):", error);
        return NextResponse.json({ error: "Failed to fetch history.", details: error.message }, { status: 500 });
    }
}
