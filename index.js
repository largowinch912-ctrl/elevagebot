const {
  Client, GatewayIntentBits, ActionRowBuilder,
  ButtonBuilder, ButtonStyle, Events,
  ModalBuilder, TextInputBuilder, TextInputStyle,
  REST, Routes, SlashCommandBuilder
} = require('discord.js');
require('dotenv').config();
const fs = require('fs');
const XLSX = require('xlsx');

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
  ]
});

// ─── Config ───────────────────────────────────────────────────────
const CHANNEL_ID = process.env.CHANNEL_ID;
const STOCK_FILE = 'stock.json';

// ─── Listes ───────────────────────────────────────────────────────
const DRAGODINDES = [
  'Amande', 'Rousse', 'Dorée',
  'Amande/Rousse', 'Amande/Dorée', 'Dorée/Rousse',
  'Ebène', 'Indigo',
  'Amande/Ebène', 'Amande/Indigo', 'Ebène/Rousse', 'Indigo/Rousse',
  'Dorée/Ebène', 'Dorée/Indigo', 'Ebène/Indigo',
  'Pourpre', 'Orchidée',
  'Amande/Pourpre', 'Amande/Orchidée', 'Pourpre/Rousse', 'Rousse/Orchidée',
  'Dorée/Pourpre', 'Dorée/Orchidée', 'Ebène/Pourpre', 'Ebène/Orchidée',
  'Indigo/Pourpre', 'Indigo/Orchidée', 'Orchidée/Pourpre',
  'Ivoire', 'Turquoise',
  'Ivoire/Amande', 'Turquoise/Amande', 'Ivoire/Rousse', 'Turquoise/Rousse',
  'Dorée/Ivoire', 'Dorée/Turquoise', 'Ebène/Ivoire', 'Ebène/Turquoise',
  'Indigo/Ivoire', 'Indigo/Turquoise', 'Ivoire/Pourpre', 'Turquoise/Pourpre',
  'Orchidée/Ivoire', 'Turquoise/Orchidée', 'Ivoire/Turquoise',
  'Prune', 'Emeraude',
  'Prune/Amande', 'Emeraude/Amande', 'Prune/Rousse', 'Emeraude/Rousse',
  'Prune/Dorée', 'Emeraude/Dorée', 'Ebène/Prune', 'Ebène/Emeraude',
  'Prune/Indigo', 'Emeraude/Indigo', 'Prune/Pourpre', 'Emeraude/Pourpre',
  'Prune/Orchidée', 'Emeraude/Orchidée', 'Prune/Ivoire', 'Emeraude/Ivoire',
  'Prune/Turquoise', 'Emeraude/Turquoise', 'Prune/Emeraude',
];

const PARCHOS = [
  'Petit Vitalité', 'Petit Force', 'Petit Intelligence',
  'Petit Chance', 'Petit Agilité', 'Petit Sagesse',
  'Parchemin Vitalité', 'Parchemin Force', 'Parchemin Intelligence',
  'Parchemin Chance', 'Parchemin Agilité', 'Parchemin Sagesse',
  'Grand Vitalité', 'Grand Force', 'Grand Intelligence',
  'Grand Chance', 'Grand Agilité', 'Grand Sagesse',
  'Puissant Vitalité', 'Puissant Force', 'Puissant Intelligence',
  'Puissant Chance', 'Puissant Agilité', 'Puissant Sagesse',
  'Parchemin Doré',
];

// ─── Stock ────────────────────────────────────────────────────────
function loadStock() {
  if (!fs.existsSync(STOCK_FILE)) return { dd: [], parchos: {} };
  return JSON.parse(fs.readFileSync(STOCK_FILE));
}

function saveStock(stock) {
  fs.writeFileSync(STOCK_FILE, JSON.stringify(stock, null, 2));
}

// ─── Trouver le meilleur match dans une liste ─────────────────────
function findMatch(input, list) {
  const q = input.toLowerCase().trim();
  // exact
  const exact = list.find(i => i.toLowerCase() === q);
  if (exact) return exact;
  // commence par
  const starts = list.find(i => i.toLowerCase().startsWith(q));
  if (starts) return starts;
  // contient
  const contains = list.find(i => i.toLowerCase().includes(q));
  if (contains) return contains;
  return null;
}

// ─── Générer Excel ────────────────────────────────────────────────
function generateExcel(stock) {
  const wb = XLSX.utils.book_new();

  // Feuille DD
  const ddRows = [['Type', 'Sexe', 'Quantité']];
  for (const entry of stock.dd) {
    ddRows.push([entry.type, entry.sexe, entry.quantite]);
  }
  const wsDd = XLSX.utils.aoa_to_sheet(ddRows);
  wsDd['!cols'] = [{ wch: 30 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsDd, 'Dragodindes');

  // Feuille Parchos
  const parchoRows = [['Type', 'Quantité']];
  for (const [type, qty] of Object.entries(stock.parchos)) {
    if (qty > 0) parchoRows.push([type, qty]);
  }
  const wsParchos = XLSX.utils.aoa_to_sheet(parchoRows);
  wsParchos['!cols'] = [{ wch: 30 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsParchos, 'Parchos');

  const path = '/tmp/stock_elevage.xlsx';
  XLSX.writeFile(wb, path);
  return path;
}

// ─── Boutons principaux ───────────────────────────────────────────
function getMainButtons() {
  return [
    new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId('dd_naissance').setLabel('🥚 Naissance DD').setStyle(ButtonStyle.Success),
      new ButtonBuilder().setCustomId('dd_vente').setLabel('💀 Vente DD').setStyle(ButtonStyle.Danger),
      new ButtonBuilder().setCustomId('parcho_gain').setLabel('📜 Gain Parcho').setStyle(ButtonStyle.Primary),
      new ButtonBuilder().setCustomId('parcho_vente').setLabel('💰 Vente Parcho').setStyle(ButtonStyle.Secondary),
    )
  ];
}

// ─── Slash command /stock ─────────────────────────────────────────
async function registerSlashCommand() {
  const rest = new REST({ version: '10' }).setToken(process.env.BOT_TOKEN);
  const commands = [
    new SlashCommandBuilder().setName('stock').setDescription('Affiche et envoie le fichier Excel du stock'),
    new SlashCommandBuilder().setName('elevage').setDescription('Affiche le panneau de gestion élevage'),
  ].map(c => c.toJSON());

  const guilds = client.guilds.cache.map(g => g.id);
  for (const guildId of guilds) {
    await rest.put(Routes.applicationGuildCommands(process.env.CLIENT_ID, guildId), { body: commands });
  }
}

// ─── Ready ────────────────────────────────────────────────────────
client.once(Events.ClientReady, async c => {
  console.log(`✅ Connecté en tant que ${c.user.tag}`);
  await registerSlashCommand();
});

// ─── Interactions ─────────────────────────────────────────────────
client.on(Events.InteractionCreate, async interaction => {

  // ── Slash /stock ──
  if (interaction.isChatInputCommand() && interaction.commandName === 'stock') {
    const stock = loadStock();
    const path = generateExcel(stock);

    // Résumé texte
    const totalDD = stock.dd.reduce((s, e) => s + e.quantite, 0);
    const totalParchos = Object.values(stock.parchos).reduce((s, v) => s + v, 0);
    let resume = `📊 **Stock actuel**\n🐔 **Dragodindes :** ${totalDD} au total\n`;
    for (const entry of stock.dd) {
      resume += `  • ${entry.type} (${entry.sexe}) : ${entry.quantite}\n`;
    }
    resume += `\n📜 **Parchos :** ${totalParchos} au total\n`;
    for (const [type, qty] of Object.entries(stock.parchos)) {
      if (qty > 0) resume += `  • ${type} : ${qty}\n`;
    }

    await interaction.reply({
      content: resume || 'Stock vide !',
      files: [{ attachment: path, name: 'stock_elevage.xlsx' }],
      ephemeral: true,
    });
    return;
  }

  // ── Slash /elevage ──
  if (interaction.isChatInputCommand() && interaction.commandName === 'elevage') {
    await interaction.reply({
      content: '🐔 **Gestion de l\'élevage**\nUtilise les boutons pour mettre à jour ton stock :',
      components: getMainButtons(),
    });
    return;
  }

  // ── Boutons → Modals ──
  if (interaction.isButton()) {
    const modals = {
      dd_naissance: { title: '🥚 Naissance de Dragodindes', fields: ['type_dd', 'sexe_dd', 'quantite'] },
      dd_vente:     { title: '💀 Vente de Dragodindes',     fields: ['type_dd', 'sexe_dd', 'quantite'] },
      parcho_gain:  { title: '📜 Gain de Parchemins',       fields: ['type_parcho', 'quantite'] },
      parcho_vente: { title: '💰 Vente de Parchemins',      fields: ['type_parcho', 'quantite'] },
    };

    const config = modals[interaction.customId];
    if (!config) return;

    const modal = new ModalBuilder()
      .setCustomId(`modal_${interaction.customId}`)
      .setTitle(config.title);

    if (config.fields.includes('type_dd')) {
      modal.addComponents(
        new ActionRowBuilder().addComponents(
          new TextInputBuilder().setCustomId('type_dd').setLabel('Type de DD (ex: Emeraude, Prune/Rousse...)').setStyle(TextInputStyle.Short).setRequired(true)
        ),
        new ActionRowBuilder().addComponents(
          new TextInputBuilder().setCustomId('sexe_dd').setLabel('Sexe (M ou F)').setStyle(TextInputStyle.Short).setMaxLength(1).setRequired(true)
        ),
        new ActionRowBuilder().addComponents(
          new TextInputBuilder().setCustomId('quantite').setLabel('Quantité').setStyle(TextInputStyle.Short).setRequired(true)
        )
      );
    } else {
      modal.addComponents(
        new ActionRowBuilder().addComponents(
          new TextInputBuilder().setCustomId('type_parcho').setLabel('Type de parcho (ex: Puissant Vitalité...)').setStyle(TextInputStyle.Short).setRequired(true)
        ),
        new ActionRowBuilder().addComponents(
          new TextInputBuilder().setCustomId('quantite').setLabel('Quantité').setStyle(TextInputStyle.Short).setRequired(true)
        )
      );
    }

    await interaction.showModal(modal);
    return;
  }

  // ── Modal submit ──
  if (interaction.isModalSubmit()) {
    const stock = loadStock();
    const action = interaction.customId.replace('modal_', '');
    const qty = parseInt(interaction.fields.getTextInputValue('quantite'));

    if (isNaN(qty) || qty <= 0) {
      return interaction.reply({ content: '❌ Quantité invalide !', ephemeral: true });
    }

    // DD
    if (action === 'dd_naissance' || action === 'dd_vente') {
      const inputType = interaction.fields.getTextInputValue('type_dd');
      const inputSexe = interaction.fields.getTextInputValue('sexe_dd').toUpperCase();

      const matchedType = findMatch(inputType, DRAGODINDES);
      if (!matchedType) {
        return interaction.reply({ content: `❌ Type de DD non reconnu : \`${inputType}\`\nEssaie un nom plus précis.`, ephemeral: true });
      }
      if (!['M', 'F'].includes(inputSexe)) {
        return interaction.reply({ content: '❌ Sexe invalide, entre M ou F.', ephemeral: true });
      }

      const existing = stock.dd.find(e => e.type === matchedType && e.sexe === inputSexe);

      if (action === 'dd_naissance') {
        if (existing) existing.quantite += qty;
        else stock.dd.push({ type: matchedType, sexe: inputSexe, quantite: qty });
        saveStock(stock);
        await interaction.reply({ content: `✅ +${qty} **${matchedType}** (${inputSexe}) ajouté${qty > 1 ? 's' : ''} au stock !`, ephemeral: true });
      } else {
        if (!existing || existing.quantite < qty) {
          return interaction.reply({ content: `❌ Stock insuffisant pour **${matchedType}** (${inputSexe}) — tu en as ${existing?.quantite ?? 0}.`, ephemeral: true });
        }
        existing.quantite -= qty;
        if (existing.quantite === 0) stock.dd = stock.dd.filter(e => !(e.type === matchedType && e.sexe === inputSexe));
        saveStock(stock);
        await interaction.reply({ content: `✅ -${qty} **${matchedType}** (${inputSexe}) vendu${qty > 1 ? 's' : ''} !`, ephemeral: true });
      }
    }

    // Parchos
    if (action === 'parcho_gain' || action === 'parcho_vente') {
      const inputParcho = interaction.fields.getTextInputValue('type_parcho');
      const matchedParcho = findMatch(inputParcho, PARCHOS);

      if (!matchedParcho) {
        return interaction.reply({ content: `❌ Type de parcho non reconnu : \`${inputParcho}\`\nEssaie un nom plus précis.`, ephemeral: true });
      }

      if (!stock.parchos[matchedParcho]) stock.parchos[matchedParcho] = 0;

      if (action === 'parcho_gain') {
        stock.parchos[matchedParcho] += qty;
        saveStock(stock);
        await interaction.reply({ content: `✅ +${qty} **${matchedParcho}** ajouté${qty > 1 ? 's' : ''} !`, ephemeral: true });
      } else {
        if (stock.parchos[matchedParcho] < qty) {
          return interaction.reply({ content: `❌ Stock insuffisant pour **${matchedParcho}** — tu en as ${stock.parchos[matchedParcho]}.`, ephemeral: true });
        }
        stock.parchos[matchedParcho] -= qty;
        saveStock(stock);
        await interaction.reply({ content: `✅ -${qty} **${matchedParcho}** vendu${qty > 1 ? 's' : ''} !`, ephemeral: true });
      }
    }
  }
});

client.login(process.env.BOT_TOKEN);