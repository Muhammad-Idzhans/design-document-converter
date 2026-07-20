import { NextResponse } from 'next/server';
import { CosmosClient } from '@azure/cosmos';

// Initialize Cosmos DB Client
const endpoint = process.env.COSMOS_DB_URI || "";
const key = process.env.COSMOS_DB_KEY || "";
const databaseId = process.env.COSMOS_DB_DATABASE || "design-doc-generator";
const containerId = process.env.COSMOS_DB_SETTINGS_CONTAINER || "settings";

let client: CosmosClient | null = null;
if (endpoint && key) {
    client = new CosmosClient({ endpoint, key });
}

export async function GET() {
    try {
        if (!client) {
            return NextResponse.json({ error: "Cosmos DB not configured" }, { status: 500 });
        }
        
        const database = client.database(databaseId);
        const container = database.container(containerId);
        
        // We will store settings under a single fixed ID "app-settings"
        const { resource: settings } = await container.item("app-settings", "app-settings").read();
        
        if (!settings) {
            // Fallback default structure if DB is brand new
            return NextResponse.json({ error: "Settings not found" }, { status: 404 });
        }
        
        return NextResponse.json(settings);
    } catch (error: any) {
        if (error.code === 404) {
            // Return default settings if DB is completely brand new
            return NextResponse.json({
                pricing_config: {
                    gpt_4o_vision: { prompt: 2.50, completion: 10.00 },
                    gpt_4_1: { prompt: 2.00, completion: 8.00 },
                    content_understanding: { rate_per_page: 5.00 }
                },
                exchange_rates: { USD_TO_MYR: 4.2 }
            });
        }
        console.error("Failed to read settings from Cosmos DB:", error);
        return NextResponse.json({ error: "Failed to read settings" }, { status: 500 });
    }
}

export async function POST(request: Request) {
    try {
        if (!client) {
            return NextResponse.json({ error: "Cosmos DB not configured" }, { status: 500 });
        }
        
        const body = await request.json();
        
        // Generate Malaysian time
        const now = new Date();
        const formatter = new Intl.DateTimeFormat('en-GB', {
            timeZone: 'Asia/Kuala_Lumpur',
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: true
        });
        
        // Format to e.g., "08/07/2026 04:30:00 pm"
        const formattedDate = formatter.format(now).replace(',', '');
        body.lastUpdated = `${formattedDate} (MYT)`;
        
        // Required for Cosmos DB
        body.id = "app-settings";

        const database = client.database(databaseId);
        const container = database.container(containerId);
        
        // Upsert to Cosmos DB
        const { resource: updatedSettings } = await container.items.upsert(body);

        return NextResponse.json({ 
            success: true, 
            message: "Settings saved to Cosmos DB successfully",
            settings: updatedSettings 
        });
    } catch (error) {
        console.error("Failed to save settings to Cosmos DB:", error);
        return NextResponse.json({ error: "Failed to save settings" }, { status: 500 });
    }
}
