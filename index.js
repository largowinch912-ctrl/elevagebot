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

const STOCK_FILE = 'stock.json';

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

function loadStock() {
  if (!fs.existsSync(STOCK_FILE)) return { dd: [], dd_vente: [], parchos: {} };
  const s = JSON.parse(fs.readFileSync(STOCK_FILE));
  if (!s.dd_vente) s.dd_vente = [];
  return s;
}

function saveStock(stock) {
  fs.writeFileSync(STOCK_FILE, JSON.stringify(stock, null, 2));
}

function findMatch(input, list) {
  const q = input.toLowerCase().trim();
  const exact = list.find(i => i.toLowerCase() === q);
  if (exact) return exact;
  const starts = list.find(i => i.toLowerCase().startsWith(q));
  if (starts) return starts;
  const contains = list.find(i => i.toLowerCase().includes(q));
  if (contains) return contains;
  return null;
}

function generateExcel(stock) {
  const wb = XLSX.utils.book_new();

  const ddRows = [['Type', 'Sexe', 'Quantité']];
  for (const e of stock.dd) ddRows.push([e.type, e.sexe, e.quantite]);
  const wsDd = XLSX.utils.aoa_to_sheet(ddRows);
  wsDd['!cols'] = [{ wch: 30 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsDd, 'Stock DD');

  const venteRows = [['Type', 'Sexe', 'Quantité']];
  for (const e of stock.dd_vente) venteRows.push([e.type, e.sexe, e.quantite]);
  const wsVente = XLSX.utils.aoa_to_sheet(venteRows);
  wsVente['!cols'] = [{ wch: 30 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsVente, 'DD à vendre');

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

function getMainButtons() {
  return [
    new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId('dd_naissance').setLabel('🥚 Naissance DD').setStyle(ButtonStyle.Success),
      new ButtonBuilder().setCustomId('dd_enlever').setLabel('❌ Enlever DD').setStyle(ButtonStyle.Danger),
    ),
    new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId('dd_avendre').setLabel('🏷️ DD à vendre').setStyle(ButtonStyle.Primary),
      new ButtonBuilder().setCustomId('dd_vendu').setLabel('💸 DD vendue').setStyle(ButtonStyle.Secondary),
    ),
    new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId('parcho_gain').setLabel('📜 Gain Parcho').setStyle(ButtonStyle.Primary),
      new ButtonBuilder().setCustomId('parcho_vente').setLabel('💰 Vente Parcho').setStyle(ButtonStyle.Secondary),
    )
  ];
}

async function registerSlashCommand() {
  const rest = new REST({ version: '10' }).setToken(process.env.BOT_TOKEN);
  const commands = [
    new SlashCommandBuilder().setName('stock').setDescription('Envoie le fichier Excel du stock'),
    new SlashCommandBuilder().setName('elevage').setDescription('Affiche le panneau de gestion élevage'),
  ].map(c => c.toJSON());

  const guilds = client.guilds.cache.map(g => g.id);
  for (const guildId of guilds) {
    await rest.put(Routes.applicationGuildCommands(process.env.CLIENT_ID, guildId), { body: commands });
  }
}

client.once(Events.ClientReady, async c => {
  console.log(`✅ Connecté en tant que ${c.user.tag}`);
  registerSlashCommand().catch(console.error);
});

client.on(Events.InteractionCreate, async interaction => {

  if (interaction.isChatInputCommand() && interaction.commandName === 'stock') {
    const stock = loadStock();
    const path = generateExcel(stock);
    const totalDD = stock.dd.reduce((s, e) => s + e.quantite, 0);
    const totalVente = stock.dd_vente.reduce((s, e) => s + e.quantite, 0);
    const totalParchos = Object.values(stock.parchos).reduce((s, v) => s + v, 0);
    const resume = `📊 **Stock actuel**\n🐔 **DD en élevage :** ${totalDD}\n🏷️ **DD à vendre :** ${totalVente}\n📜 **Parchos :** ${totalParchos}`;
    await interaction.reply({
      content: resume,
      files: [{ attachment: path, name: 'stock_elevage.xlsx' }],
      ephemeral: true,
    });
    return;
  }

  if (interaction.isChatInputCommand() && interaction.commandName === 'elevage') {
    await interaction.reply({
      content: '🐔 **Gestion de l\'élevage**',
      components: getMainButtons(),
    });
    return;
  }

  if (interaction.isButton()) {
    const titles = {
      dd_naissance: '🥚 Naissance de Dragodindes',
      dd_avendre:   '🏷️ DD à mettre en vente',
      dd_enlever:   '❌ Enlever des Dragodindes',
      parcho_gain:  '📜 Gain de Parchemins',
      parcho_vente: '💰 Vente de Parchemins',
      dd_vendu: '💸 DD vendue',
    };
    if (!titles[interaction.customId]) return;

    const modal = new ModalBuilder()
      .setCustomId(`modal_${interaction.customId}`)
      .setTitle(titles[interaction.customId]);

    if (['dd_naissance', 'dd_avendre', 'dd_enlever', 'dd_vendu'].includes(interaction.customId)) {
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

  if (interaction.isModalSubmit()) {
    const stock = loadStock();
    const action = interaction.customId.replace('modal_', '');
    const qty = parseInt(interaction.fields.getTextInputValue('quantite'));

    if (isNaN(qty) || qty <= 0) {
      return interaction.reply({ content: '❌ Quantité invalide !', ephemeral: true });
    }

    if (['dd_naissance', 'dd_avendre', 'dd_enlever', 'dd_vendu'].includes(action)) {
      const inputType = interaction.fields.getTextInputValue('type_dd');
      const inputSexe = interaction.fields.getTextInputValue('sexe_dd').toUpperCase();
      const matchedType = findMatch(inputType, DRAGODINDES);

      if (!matchedType) return interaction.reply({ content: `❌ Type non reconnu : \`${inputType}\``, ephemeral: true });
      if (!['M', 'F'].includes(inputSexe)) return interaction.reply({ content: '❌ Sexe invalide, entre M ou F.', ephemeral: true });

      if (action === 'dd_naissance') {
        const existing = stock.dd.find(e => e.type === matchedType && e.sexe === inputSexe);
        if (existing) existing.quantite += qty;
        else stock.dd.push({ type: matchedType, sexe: inputSexe, quantite: qty });
        saveStock(stock);
        return interaction.reply({ content: `✅ +${qty} **${matchedType}** (${inputSexe}) ajouté au stock !`, ephemeral: true });
      }

      if (action === 'dd_avendre') {
        const existing = stock.dd_vente.find(e => e.type === matchedType && e.sexe === inputSexe);
        if (existing) existing.quantite += qty;
        else stock.dd_vente.push({ type: matchedType, sexe: inputSexe, quantite: qty });
        saveStock(stock);
        return interaction.reply({ content: `✅ +${qty} **${matchedType}** (${inputSexe}) ajouté aux DD à vendre !`, ephemeral: true });
      }

      if (action === 'dd_enlever') {
        const existing = stock.dd.find(e => e.type === matchedType && e.sexe === inputSexe);
        if (!existing || existing.quantite < qty) {
          return interaction.reply({ content: `❌ Stock insuffisant : tu as ${existing?.quantite ?? 0} **${matchedType}** (${inputSexe}).`, ephemeral: true });
        }
        existing.quantite -= qty;
        if (existing.quantite === 0) stock.dd = stock.dd.filter(e => !(e.type === matchedType && e.sexe === inputSexe));
        saveStock(stock);
        return interaction.reply({ content: `✅ -${qty} **${matchedType}** (${inputSexe}) retiré du stock.`, ephemeral: true });
      }

      if (action === 'dd_vendu') {
        const existing = stock.dd_vente.find(e => e.type === matchedType && e.sexe === inputSexe);
        if (!existing || existing.quantite < qty) {
          return interaction.reply({ content: `❌ Stock insuffisant dans "à vendre" : tu as ${existing?.quantite ?? 0} **${matchedType}** (${inputSexe}).`, ephemeral: true });
        }
        existing.quantite -= qty;
        if (existing.quantite === 0) stock.dd_vente = stock.dd_vente.filter(e => !(e.type === matchedType && e.sexe === inputSexe));
        saveStock(stock);
        return interaction.reply({ content: `✅ -${qty} **${matchedType}** (${inputSexe}) vendu depuis le stock à vendre !`, ephemeral: true });
      }
    }

    if (['parcho_gain', 'parcho_vente'].includes(action)) {
      const inputParcho = interaction.fields.getTextInputValue('type_parcho');
      const matchedParcho = findMatch(inputParcho, PARCHOS);
      if (!matchedParcho) return interaction.reply({ content: `❌ Parcho non reconnu : \`${inputParcho}\``, ephemeral: true });

      if (!stock.parchos[matchedParcho]) stock.parchos[matchedParcho] = 0;

      if (action === 'parcho_gain') {
        stock.parchos[matchedParcho] += qty;
        saveStock(stock);
        return interaction.reply({ content: `✅ +${qty} **${matchedParcho}** ajouté !`, ephemeral: true });
      } else {
        if (stock.parchos[matchedParcho] < qty) {
          return interaction.reply({ content: `❌ Stock insuffisant : tu as ${stock.parchos[matchedParcho]} **${matchedParcho}**.`, ephemeral: true });
        }
        stock.parchos[matchedParcho] -= qty;
        saveStock(stock);
        return interaction.reply({ content: `✅ -${qty} **${matchedParcho}** vendu !`, ephemeral: true });
      }
    }
  }
});

client.login(process.env.BOT_TOKEN);