// Example of using IndexedDB
const dbName = "TestDatabase";
const storeName = "TestStore";

// Open (or create) the database
const request = indexedDB.open(dbName, 1);

request.onupgradeneeded = function (event) {
    const db = event.target.result;

    // Create an object store if it doesn't exist
    if (!db.objectStoreNames.contains(storeName)) {
        db.createObjectStore(storeName, { keyPath: "id", autoIncrement: true });
    }
};

request.onsuccess = function (event) {

    const db = event.target.result;

    // Add data to the object store
    const transaction = db.transaction(storeName, "readwrite");
    const store = transaction.objectStore(storeName);

    const data = { name: "Example", value: 42 };
    store.add(data);

    transaction.oncomplete = function () {
        console.log("Data added successfully!");

        // ✅ DELETE the user after adding
        const deleteTransaction = db.transaction(["users"], "readwrite");
        const deleteStore = deleteTransaction.objectStore("users");

        const deleteRequest = deleteStore.delete(2); // ID of the user to delete

        deleteRequest.onsuccess = function () {
            console.log("User deleted (ID: 2)");
        };

        deleteRequest.onerror = function () {
            console.error("Failed to delete user");
        };
    };

    transaction.onerror = function () {
        console.error("Error adding data:", transaction.error);
    };
};

request.onerror = function (event) {
    console.error("Error opening database:", event.target.error);
};