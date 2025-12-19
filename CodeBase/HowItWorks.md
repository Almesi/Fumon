# Fumon Game

An Excel-based game engine written in VBA with optional multiplayer support using SharePoint as a shared data layer.

> ⚠️ **Status:** Multiplayer is experimental. Current development focus is singleplayer stability.

---

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
  - [Multiplayer Concept](#multiplayer-concept)
  - [File Structure](#file-structure)
  - [Shared State via Ranges](#shared-state-via-ranges)
- [Core Systems](#core-systems)
  - [GameServer](#gameserver)
  - [GameMap](#gamemap)
  - [Players](#players)
- [Fumon System](#fumon-system)
  - [Fumon](#fumon)
  - [Fumons](#fumons)
  - [FumonDefinition](#fumondefinition)
  - [FumonBox](#fumonbox)
- [Combat System](#combat-system)
  - [Attack](#attack)
  - [Fight](#fight)
- [Items](#items)
- [Scripts](#scripts)
- [Notes](#notes)

---

## Overview

This project implements a **turn-based RPG-style game** entirely in Excel using VBA.  
Because VBA cannot share memory between users, multiplayer functionality is achieved by sharing Excel cell values through a **SharePoint-hosted workbook**.

All game logic runs locally, while shared state is synchronized through worksheet cells.

---

## Architecture

### Multiplayer Concept

- VBA cannot share variables between users
- Excel files hosted on SharePoint allow multiple users to edit the same cells
- Macros cannot run in SharePoint Excel
- All logic executes locally
- Only worksheet values are shared between players

This effectively turns Excel cells into a **networked data store**.

---

### File Structure

| File | Description |
|----|----|
| `Player.xlsm` | Local game client. Each player has their own copy. |
| `Server.xlsm` | Shared game state. Local for singleplayer, SharePoint-hosted for multiplayer. |
| `CodeBase` | Dynamically loaded VBA code to ensure version consistency. |

---

### Shared State via Ranges

#### `Range`

The `Range` object acts as a **shared pointer**.

- `Range.Value` can be read and written by all players
- Any class holding a `Range` represents shared game state

#### `IRange`

`IRange` is an interface abstraction over `Range` that:

- Reduces read/write operations
- Improves performance
- Allows optimized implementations

---

## Core Systems

### GameServer

`GameServer` is the global access point for all shared data.

- Created statically
- One instance per workbook
- Ensures all players reference identical cell addresses

This guarantees consistent shared state across all clients.

---

### GameMap

Manages all map-related data and logic.

**Responsibilities:**

- Sprite folder
- Game time
- Spawn positions
- Map grid

**Map Details:**

- Sheet name: `*Map`
- Two-dimensional grid
- Size determined by `Rows.Value` and `Columns.Value`
- `MapPointer` defines the first map cell
- `Tiles` contains tile definitions (not map cells)

---

### Players

All player types are created through the `IPlayer` interface.

#### Player Types

- **HumanPlayer** – Controlled by a real user
- **ComPlayer** – Static NPCs
- **WildPlayer** – Temporary combat-only entities

**Notes:**

- HumanPlayer and ComPlayer share most logic
- HumanPlayers can only be moved manually
- WildPlayers have no money, quests, or free movement
- All players include AI for combat decisions

---

## Fumon System

### Fumon

Fumons are the core combat entities.

**Rules:**

- Max 8 Fumons per player
- Up to 4 attacks
- Up to 2 types
- Max level: 100 (`Long`)

**Stats (defined at level 100):**

- Health
- Attack
- Defense
- Special Attack
- Special Defense
- Initiative

**Additional Features:**

- Status effects (burned, frozen, etc.)
- Evolution via level-based conditions

---

### Fumons

Container class for managing up to 8 `Fumon` objects, with helper methods for batch operations.

---

### FumonDefinition

Defines immutable Fumon data:

- ID
- Name
- Types
- Base stats
- Learnable attacks

`FumonDefinition` acts as a **factory**.  
`Fumon` instances calculate dynamic values based on these definitions.

---

### FumonBox

Stores overflow Fumons beyond the 8-slot limit.

- Stores definition ID and level only
- Switching resets experience
- Exists only in code
- Switching logic not yet implemented

---

## Combat System

### Attack

Defines how attacks modify Fumon stats.

Attack behavior depends on:

- Attacker type
- Target type
- Attack type
- Stats used for calculation
- Base power

All attack definitions are stored in the `Attacks` worksheet.

---

### Fight

Manages combat interactions using shared pointers.

**Flow:**

1. Player initiates fight
2. Check if opponent is already fighting
3. Allocate first free cell in `Fights` sheet
4. Create Fight object
5. Enter fight loop

**Turn System:**

- Each turn lasts 60 seconds
- Player actions are written to `IPlayer`
- No action results in skipped turn
- AI controls non-human players

**End Conditions:**

- Winner stored in `Winner.Value`
- Player wins if at least one Fumon remains alive
- Destroyed pointers indicate fight completion or loss

---

## Items

### Item

Items consist of:

- Pointer to an `ItemDefinition`
- Pointer to owned quantity

**Behavior:**

- Using an item reduces quantity by 1
- Some items are non-usable (progression only)
- Item effects are defined via scripts

---

## Scripts

Scripts turn Excel into a **game engine**.

### Script Use Cases

- Item behavior
- Fumon evolution conditions
- Game win conditions
- Fight initiation
- UI messaging
- Player and NPC movement
- WildPlayer spawning
- NPC sight detection

### Script Fields

| Field | Description |
|----|----|
| `Index` | Script identifier |
| `Name` | Script name |
| `Text` | Callable VBA string |
| `ExecuteTimer` | In-game minutes between executions (0 = every frame) |

---

## Notes

- Multiplayer relies entirely on shared worksheet values
- Pointer destruction is used as a synchronization mechanism
- Performance depends heavily on minimizing `Range.Value` access
- Performance depends heavily on minimizing `std_Callable` functions

---

## License

Not specified.

